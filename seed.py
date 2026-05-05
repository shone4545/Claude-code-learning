"""Seed the database with realistic sample batches for demo / smoke testing.

Mirrors the user's operation: 4 intakes per week (Mon/Tue/Thu/Fri),
~575k chicks each, staggered. Generates daily records for in-flight batches
so the dashboards have meaningful charts on first load.
"""
from __future__ import annotations

import math
import random
from datetime import date, timedelta

from poultry import create_app
from poultry.extensions import db
from poultry.metrics import (
    DEFAULT_VACCINE_SCHEDULE,
    COBB500_BODY_WEIGHT_G,
    standard_weight_for_age,
)
from poultry.models import (
    Batch,
    DailyRecord,
    HarvestEvent,
    ManpowerLog,
    UtilityReading,
    VaccineEvent,
)


INTAKE_WEEKDAYS = [0, 1, 3, 4]  # Mon, Tue, Thu, Fri
HOUSES = ["House 1", "House 2", "House 3", "House 4", "House 5", "House 6"]


def _intake_dates(weeks_back: int, today: date) -> list[date]:
    dates: list[date] = []
    start = today - timedelta(days=weeks_back * 7)
    for offset in range(weeks_back * 7 + 1):
        d = start + timedelta(days=offset)
        if d.weekday() in INTAKE_WEEKDAYS and d <= today:
            dates.append(d)
    return dates


def _daily_feed_kg(age_days: int, birds: int) -> float:
    # Approx Cobb 500 daily feed intake per bird (g), then scale.
    table = {
        0: 13, 3: 25, 7: 40, 10: 55, 14: 75, 17: 95, 21: 115,
        24: 135, 28: 155, 31: 170, 35: 185, 38: 195, 42: 205,
    }
    keys = sorted(table)
    if age_days <= keys[0]:
        per_bird_g = table[keys[0]]
    elif age_days >= keys[-1]:
        per_bird_g = table[keys[-1]]
    else:
        for k0, k1 in zip(keys, keys[1:]):
            if k0 <= age_days <= k1:
                t = (age_days - k0) / (k1 - k0)
                per_bird_g = table[k0] + t * (table[k1] - table[k0])
                break
    return birds * per_bird_g / 1000.0


def seed():
    today = date.today()
    rng = random.Random(42)

    db.drop_all()
    db.create_all()

    intake_dates = _intake_dates(weeks_back=6, today=today)
    print(f"Seeding {len(intake_dates)} batches...")

    for i, intake in enumerate(intake_dates):
        intake_count = rng.randint(550_000, 600_000)
        house = HOUSES[i % len(HOUSES)]
        code = f"B{intake.strftime('%Y%m%d')}-{house.split()[-1]}"
        batch = Batch(
            batch_code=code,
            intake_date=intake,
            intake_count=intake_count,
            breed="Cobb 500",
            house=house,
            target_weight_min_g=2100,
            target_weight_max_g=2500,
            status="active",
        )
        db.session.add(batch)
        db.session.flush()

        for entry in DEFAULT_VACCINE_SCHEDULE:
            sched = intake + timedelta(days=entry["age"])
            ev = VaccineEvent(
                batch_id=batch.id,
                scheduled_date=sched,
                age_days=entry["age"],
                vaccine_name=entry["name"],
                route=entry["route"],
                doses_planned=intake_count,
            )
            if sched <= today:
                ev.administered_date = sched
                ev.doses_used = int(intake_count * rng.uniform(0.95, 1.02))
            db.session.add(ev)

        # Daily records up to today.
        alive = intake_count
        days_alive = (today - intake).days
        # Mortality bias: higher first week, then declining.
        for d in range(0, days_alive + 1):
            rec_date = intake + timedelta(days=d)
            base_mort_pct = 0.0008 if d < 7 else 0.0003
            mortality = max(0, int(alive * rng.uniform(base_mort_pct * 0.5, base_mort_pct * 1.8)))
            culls = rng.randint(0, max(1, mortality // 3))
            alive -= mortality + culls

            feed = _daily_feed_kg(d, alive) * rng.uniform(0.95, 1.05)
            water = feed * rng.uniform(1.6, 2.0)
            std_weight = standard_weight_for_age(d) or 0
            actual_weight = std_weight * rng.uniform(0.94, 1.04) if std_weight else None
            # Sample weight only twice a week.
            if d % 3 != 0:
                actual_weight = None

            db.session.add(
                DailyRecord(
                    batch_id=batch.id,
                    record_date=rec_date,
                    mortality_count=mortality,
                    cull_count=culls,
                    feed_consumed_kg=round(feed, 1),
                    water_consumed_l=round(water, 1),
                    avg_body_weight_g=round(actual_weight, 0) if actual_weight else None,
                    sample_size=100 if actual_weight else None,
                    temperature_c=round(rng.uniform(28, 32) - min(d * 0.1, 4), 1),
                    humidity_pct=round(rng.uniform(60, 75), 1),
                )
            )

        # Harvest mature batches (age >= 38) and close them.
        if days_alive >= 38:
            harvest_date = intake + timedelta(days=rng.randint(38, 42))
            if harvest_date <= today:
                avg_g = rng.randint(2150, 2450)
                birds = int(alive * rng.uniform(0.96, 0.99))
                db.session.add(
                    HarvestEvent(
                        batch_id=batch.id,
                        harvest_date=harvest_date,
                        birds_harvested=birds,
                        total_weight_kg=round(birds * avg_g / 1000.0, 0),
                        avg_weight_g=avg_g,
                    )
                )
                batch.status = "harvested"

    # Manpower entries (last 14 days, ~12 workers, mixed roles).
    workers = [
        ("Ahmad", "supervisor"), ("Siti", "farmhand"), ("Kumar", "farmhand"),
        ("Rahim", "farmhand"), ("Mei Ling", "vet tech"), ("Hassan", "feed crew"),
        ("Nadia", "farmhand"), ("Tan", "feed crew"), ("Faizal", "farmhand"),
        ("Lim", "farmhand"), ("Aziz", "supervisor"), ("Ravi", "farmhand"),
    ]
    active_batches = (
        db.session.query(Batch).filter(Batch.status == "active").all()
    )
    for d_off in range(14):
        log_date = today - timedelta(days=d_off)
        for name, role in workers:
            if rng.random() > 0.85:
                continue
            db.session.add(
                ManpowerLog(
                    log_date=log_date,
                    batch_id=rng.choice(active_batches).id if active_batches and rng.random() > 0.3 else None,
                    worker_name=name,
                    role=role,
                    hours=rng.choice([8, 8, 8, 9, 10, 6]),
                    task=rng.choice(["feeding", "cleaning", "vaccination", "weighing", "biosecurity"]),
                )
            )

    # Utility readings (weekly cadence, last 8 weeks).
    for w in range(8):
        d = today - timedelta(days=w * 7)
        db.session.add(UtilityReading(reading_date=d, utility_type="electricity",
                                      quantity=round(rng.uniform(8000, 12000), 0), unit="kWh",
                                      cost=round(rng.uniform(3500, 5500), 2)))
        db.session.add(UtilityReading(reading_date=d, utility_type="water",
                                      quantity=round(rng.uniform(180000, 240000), 0), unit="L",
                                      cost=round(rng.uniform(800, 1200), 2)))
        db.session.add(UtilityReading(reading_date=d, utility_type="gas",
                                      quantity=round(rng.uniform(800, 1500), 0), unit="kg",
                                      cost=round(rng.uniform(2400, 4000), 2)))

    db.session.commit()
    print(
        f"Done. Batches: {db.session.query(Batch).count()} "
        f"(active: {db.session.query(Batch).filter_by(status='active').count()})"
    )


if __name__ == "__main__":
    app = create_app()
    with app.app_context():
        seed()
