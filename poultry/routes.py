from __future__ import annotations

from datetime import date, datetime, timedelta

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    url_for,
    abort,
)
from sqlalchemy import select, func

from .extensions import db
from .metrics import (
    DEFAULT_VACCINE_SCHEDULE,
    compute_batch_metrics,
    standard_weight_for_age,
)
from .models import (
    Batch,
    DailyRecord,
    HarvestEvent,
    ManpowerLog,
    UtilityReading,
    VaccineEvent,
)


def _parse_date(value: str | None) -> date | None:
    if not value:
        return None
    return datetime.strptime(value, "%Y-%m-%d").date()


def _parse_int(value: str | None, default: int = 0) -> int:
    if value is None or value == "":
        return default
    return int(value)


def _parse_float(value: str | None, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    return float(value)


def _parse_optional_float(value: str | None) -> float | None:
    if value is None or value == "":
        return None
    return float(value)


def _parse_optional_int(value: str | None) -> int | None:
    if value is None or value == "":
        return None
    return int(value)


def _get_batch_or_404(batch_id: int) -> Batch:
    batch = db.session.get(Batch, batch_id)
    if not batch:
        abort(404)
    return batch


def register_routes(app: Flask) -> None:
    @app.route("/")
    def dashboard():
        active = (
            db.session.execute(
                select(Batch).where(Batch.status == "active").order_by(Batch.intake_date)
            )
            .scalars()
            .all()
        )
        metrics = [compute_batch_metrics(b) for b in active]

        total_birds = sum(m.birds_alive for m in metrics)
        total_intake = sum(m.intake_count for m in metrics)
        weighted_mortality = (
            sum(m.mortality_pct * m.intake_count for m in metrics) / total_intake
            if total_intake
            else 0.0
        )
        upcoming_harvests = sorted(
            [m for m in metrics if m.projected_harvest_date],
            key=lambda m: m.projected_harvest_date,
        )[:5]

        upcoming_vaccines = (
            db.session.execute(
                select(VaccineEvent)
                .where(VaccineEvent.administered_date.is_(None))
                .where(VaccineEvent.scheduled_date <= date.today() + timedelta(days=7))
                .order_by(VaccineEvent.scheduled_date)
            )
            .scalars()
            .all()
        )

        return render_template(
            "dashboard.html",
            active_nav="dashboard",
            metrics=metrics,
            total_birds=total_birds,
            total_intake=total_intake,
            weighted_mortality=weighted_mortality,
            upcoming_harvests=upcoming_harvests,
            upcoming_vaccines=upcoming_vaccines,
        )

    # ---- Batches ----------------------------------------------------------
    @app.route("/batches")
    def list_batches():
        status = request.args.get("status", "active")
        q = select(Batch).order_by(Batch.intake_date.desc())
        if status != "all":
            q = q.where(Batch.status == status)
        batches = db.session.execute(q).scalars().all()
        metrics = {b.id: compute_batch_metrics(b) for b in batches}
        return render_template(
            "batches/list.html",
            active_nav="batches",
            batches=batches,
            metrics=metrics,
            status=status,
        )

    @app.route("/batches/new", methods=["GET", "POST"])
    def new_batch():
        if request.method == "POST":
            intake_date = _parse_date(request.form.get("intake_date")) or date.today()
            batch = Batch(
                batch_code=request.form["batch_code"].strip(),
                intake_date=intake_date,
                intake_count=_parse_int(request.form.get("intake_count")),
                breed=request.form.get("breed", "Cobb 500").strip() or "Cobb 500",
                house=request.form.get("house", "").strip(),
                target_weight_min_g=_parse_int(request.form.get("target_weight_min_g"), 2100),
                target_weight_max_g=_parse_int(request.form.get("target_weight_max_g"), 2500),
                notes=request.form.get("notes", "").strip(),
            )
            db.session.add(batch)
            db.session.flush()

            if request.form.get("auto_vaccine") == "on":
                for entry in DEFAULT_VACCINE_SCHEDULE:
                    db.session.add(
                        VaccineEvent(
                            batch_id=batch.id,
                            scheduled_date=intake_date + timedelta(days=entry["age"]),
                            age_days=entry["age"],
                            vaccine_name=entry["name"],
                            route=entry["route"],
                            doses_planned=batch.intake_count,
                        )
                    )

            db.session.commit()
            flash(f"Batch {batch.batch_code} created.", "success")
            return redirect(url_for("batch_detail", batch_id=batch.id))

        suggested_code = f"B{date.today().strftime('%Y%m%d')}"
        return render_template(
            "batches/new.html",
            active_nav="batches",
            today=date.today().isoformat(),
            suggested_code=suggested_code,
        )

    @app.route("/batches/<int:batch_id>")
    def batch_detail(batch_id):
        batch = _get_batch_or_404(batch_id)
        metrics = compute_batch_metrics(batch)

        weight_series = [
            {
                "age": (r.record_date - batch.intake_date).days,
                "date": r.record_date.isoformat(),
                "actual": r.avg_body_weight_g,
                "standard": standard_weight_for_age(
                    (r.record_date - batch.intake_date).days
                ),
            }
            for r in batch.daily_records
            if r.avg_body_weight_g
        ]
        # Pad standard curve up to projected harvest age for nicer chart.
        max_age = max(
            (m.projected_harvest_age or metrics.age_days for m in [metrics]),
            default=metrics.age_days,
        )
        std_curve = [
            {"age": a, "standard": standard_weight_for_age(a)}
            for a in range(0, max_age + 2)
        ]

        mortality_series = []
        cumulative = 0
        for r in batch.daily_records:
            cumulative += r.mortality_count + r.cull_count
            mortality_series.append(
                {
                    "age": (r.record_date - batch.intake_date).days,
                    "date": r.record_date.isoformat(),
                    "daily": r.mortality_count + r.cull_count,
                    "cumulative_pct": (
                        100.0 * cumulative / batch.intake_count
                        if batch.intake_count
                        else 0.0
                    ),
                }
            )

        return render_template(
            "batches/detail.html",
            active_nav="batches",
            batch=batch,
            m=metrics,
            weight_series=weight_series,
            std_curve=std_curve,
            mortality_series=mortality_series,
            today=date.today().isoformat(),
        )

    @app.route("/batches/<int:batch_id>/close", methods=["POST"])
    def close_batch(batch_id):
        batch = _get_batch_or_404(batch_id)
        batch.status = request.form.get("status", "closed")
        db.session.commit()
        flash(f"Batch {batch.batch_code} marked {batch.status}.", "success")
        return redirect(url_for("batch_detail", batch_id=batch.id))

    # ---- Daily records ----------------------------------------------------
    @app.route("/batches/<int:batch_id>/daily/new", methods=["GET", "POST"])
    def new_daily_record(batch_id):
        batch = _get_batch_or_404(batch_id)
        if request.method == "POST":
            record_date = _parse_date(request.form.get("record_date")) or date.today()
            existing = db.session.execute(
                select(DailyRecord).where(
                    DailyRecord.batch_id == batch.id,
                    DailyRecord.record_date == record_date,
                )
            ).scalar_one_or_none()
            record = existing or DailyRecord(batch_id=batch.id, record_date=record_date)
            record.mortality_count = _parse_int(request.form.get("mortality_count"))
            record.cull_count = _parse_int(request.form.get("cull_count"))
            record.feed_consumed_kg = _parse_float(request.form.get("feed_consumed_kg"))
            record.water_consumed_l = _parse_float(request.form.get("water_consumed_l"))
            record.avg_body_weight_g = _parse_optional_float(
                request.form.get("avg_body_weight_g")
            )
            record.sample_size = _parse_optional_int(request.form.get("sample_size"))
            record.temperature_c = _parse_optional_float(request.form.get("temperature_c"))
            record.humidity_pct = _parse_optional_float(request.form.get("humidity_pct"))
            record.notes = request.form.get("notes", "").strip()
            if not existing:
                db.session.add(record)
            db.session.commit()
            flash("Daily record saved.", "success")
            return redirect(url_for("batch_detail", batch_id=batch.id))

        prefill_date = _parse_date(request.args.get("date")) or date.today()
        existing = db.session.execute(
            select(DailyRecord).where(
                DailyRecord.batch_id == batch.id,
                DailyRecord.record_date == prefill_date,
            )
        ).scalar_one_or_none()
        return render_template(
            "batches/daily_form.html",
            active_nav="batches",
            batch=batch,
            record=existing,
            prefill_date=prefill_date.isoformat(),
            age_days=(prefill_date - batch.intake_date).days,
        )

    # ---- Vaccines ---------------------------------------------------------
    @app.route("/batches/<int:batch_id>/vaccine/new", methods=["GET", "POST"])
    def new_vaccine(batch_id):
        batch = _get_batch_or_404(batch_id)
        if request.method == "POST":
            scheduled_date = _parse_date(request.form.get("scheduled_date")) or date.today()
            ev = VaccineEvent(
                batch_id=batch.id,
                scheduled_date=scheduled_date,
                age_days=(scheduled_date - batch.intake_date).days,
                vaccine_name=request.form["vaccine_name"].strip(),
                route=request.form.get("route", "drinking water").strip(),
                doses_planned=_parse_int(request.form.get("doses_planned")),
                notes=request.form.get("notes", "").strip(),
            )
            db.session.add(ev)
            db.session.commit()
            flash("Vaccine scheduled.", "success")
            return redirect(url_for("batch_detail", batch_id=batch.id))
        return render_template(
            "batches/vaccine_form.html",
            active_nav="batches",
            batch=batch,
            today=date.today().isoformat(),
        )

    @app.route("/vaccines/<int:event_id>/administer", methods=["POST"])
    def administer_vaccine(event_id):
        ev = db.session.get(VaccineEvent, event_id)
        if not ev:
            abort(404)
        ev.administered_date = (
            _parse_date(request.form.get("administered_date")) or date.today()
        )
        ev.doses_used = _parse_optional_int(request.form.get("doses_used"))
        ev.notes = request.form.get("notes", ev.notes)
        db.session.commit()
        flash(f"Marked {ev.vaccine_name} administered.", "success")
        return redirect(url_for("batch_detail", batch_id=ev.batch_id))

    # ---- Harvest ----------------------------------------------------------
    @app.route("/harvest")
    def harvest_planner():
        active = (
            db.session.execute(
                select(Batch).where(Batch.status == "active").order_by(Batch.intake_date)
            )
            .scalars()
            .all()
        )
        rows = [compute_batch_metrics(b) for b in active]
        rows.sort(key=lambda r: r.projected_harvest_date or date.max)
        return render_template(
            "harvest_planner.html", active_nav="harvest", rows=rows
        )

    @app.route("/batches/<int:batch_id>/harvest/new", methods=["GET", "POST"])
    def new_harvest(batch_id):
        batch = _get_batch_or_404(batch_id)
        if request.method == "POST":
            harvest_date = _parse_date(request.form.get("harvest_date")) or date.today()
            birds = _parse_int(request.form.get("birds_harvested"))
            total_kg = _parse_float(request.form.get("total_weight_kg"))
            avg_g = (total_kg * 1000.0 / birds) if birds else None
            ev = HarvestEvent(
                batch_id=batch.id,
                harvest_date=harvest_date,
                birds_harvested=birds,
                total_weight_kg=total_kg,
                avg_weight_g=avg_g,
                notes=request.form.get("notes", "").strip(),
            )
            db.session.add(ev)
            db.session.commit()
            flash(f"Harvested {birds:,} birds ({total_kg:,.0f} kg).", "success")
            return redirect(url_for("batch_detail", batch_id=batch.id))
        return render_template(
            "batches/harvest_form.html",
            active_nav="batches",
            batch=batch,
            today=date.today().isoformat(),
        )

    # ---- Manpower ---------------------------------------------------------
    @app.route("/manpower")
    def list_manpower():
        days = int(request.args.get("days", 14))
        cutoff = date.today() - timedelta(days=days)
        logs = (
            db.session.execute(
                select(ManpowerLog)
                .where(ManpowerLog.log_date >= cutoff)
                .order_by(ManpowerLog.log_date.desc())
            )
            .scalars()
            .all()
        )
        total_hours = sum(l.hours for l in logs)
        return render_template(
            "manpower/list.html",
            active_nav="manpower",
            logs=logs,
            days=days,
            total_hours=total_hours,
        )

    @app.route("/manpower/new", methods=["GET", "POST"])
    def new_manpower():
        if request.method == "POST":
            log = ManpowerLog(
                log_date=_parse_date(request.form.get("log_date")) or date.today(),
                batch_id=_parse_optional_int(request.form.get("batch_id")),
                worker_name=request.form["worker_name"].strip(),
                role=request.form.get("role", "farmhand").strip(),
                hours=_parse_float(request.form.get("hours")),
                task=request.form.get("task", "").strip(),
                notes=request.form.get("notes", "").strip(),
            )
            db.session.add(log)
            db.session.commit()
            flash("Manpower entry logged.", "success")
            return redirect(url_for("list_manpower"))
        batches = (
            db.session.execute(
                select(Batch).where(Batch.status == "active").order_by(Batch.intake_date)
            )
            .scalars()
            .all()
        )
        return render_template(
            "manpower/form.html",
            active_nav="manpower",
            batches=batches,
            today=date.today().isoformat(),
        )

    # ---- Utilities --------------------------------------------------------
    @app.route("/utilities")
    def list_utilities():
        days = int(request.args.get("days", 30))
        cutoff = date.today() - timedelta(days=days)
        rows = (
            db.session.execute(
                select(UtilityReading)
                .where(UtilityReading.reading_date >= cutoff)
                .order_by(UtilityReading.reading_date.desc())
            )
            .scalars()
            .all()
        )
        totals_by_type: dict[str, dict[str, float]] = {}
        for r in rows:
            agg = totals_by_type.setdefault(
                r.utility_type, {"qty": 0.0, "cost": 0.0, "unit": r.unit}
            )
            agg["qty"] += r.quantity
            agg["cost"] += r.cost
        return render_template(
            "utilities/list.html",
            active_nav="utilities",
            rows=rows,
            days=days,
            totals_by_type=totals_by_type,
        )

    @app.route("/utilities/new", methods=["GET", "POST"])
    def new_utility():
        if request.method == "POST":
            reading = UtilityReading(
                reading_date=_parse_date(request.form.get("reading_date")) or date.today(),
                batch_id=_parse_optional_int(request.form.get("batch_id")),
                utility_type=request.form["utility_type"].strip(),
                quantity=_parse_float(request.form.get("quantity")),
                unit=request.form.get("unit", "kWh").strip() or "kWh",
                cost=_parse_float(request.form.get("cost")),
                notes=request.form.get("notes", "").strip(),
            )
            db.session.add(reading)
            db.session.commit()
            flash("Utility reading logged.", "success")
            return redirect(url_for("list_utilities"))
        batches = (
            db.session.execute(
                select(Batch).where(Batch.status == "active").order_by(Batch.intake_date)
            )
            .scalars()
            .all()
        )
        return render_template(
            "utilities/form.html",
            active_nav="utilities",
            batches=batches,
            today=date.today().isoformat(),
        )

    # Jinja filters
    @app.template_filter("comma")
    def comma_filter(value):
        if value is None:
            return "—"
        if isinstance(value, float):
            return f"{value:,.0f}"
        return f"{value:,}"

    @app.template_filter("num")
    def num_filter(value, decimals=2):
        if value is None:
            return "—"
        return f"{value:,.{decimals}f}"

    @app.template_filter("pct")
    def pct_filter(value, decimals=2):
        if value is None:
            return "—"
        return f"{value:,.{decimals}f}%"
