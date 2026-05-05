from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Iterable

from .models import Batch, DailyRecord, HarvestEvent


# Cobb 500 broiler reference body weight by age (grams). Used as fallback curve
# and as the standard for variance-from-target calculations.
COBB500_BODY_WEIGHT_G = {
    0: 42, 1: 56, 2: 72, 3: 89, 4: 109, 5: 131, 6: 156, 7: 185,
    8: 217, 9: 252, 10: 291, 11: 333, 12: 379, 13: 429, 14: 482,
    15: 540, 16: 601, 17: 666, 18: 735, 19: 808, 20: 884, 21: 964,
    22: 1048, 23: 1136, 24: 1226, 25: 1320, 26: 1417, 27: 1517,
    28: 1620, 29: 1725, 30: 1833, 31: 1944, 32: 2056, 33: 2171,
    34: 2287, 35: 2404, 36: 2523, 37: 2643, 38: 2764, 39: 2885,
    40: 3007, 41: 3129, 42: 3250,
}


def standard_weight_for_age(age_days: int) -> float | None:
    if age_days < 0:
        return None
    if age_days in COBB500_BODY_WEIGHT_G:
        return float(COBB500_BODY_WEIGHT_G[age_days])
    max_age = max(COBB500_BODY_WEIGHT_G)
    if age_days > max_age:
        return float(COBB500_BODY_WEIGHT_G[max_age])
    return None


@dataclass
class BatchMetrics:
    batch_id: int
    batch_code: str
    intake_date: date
    intake_count: int
    age_days: int
    birds_alive: int
    cumulative_mortality: int
    cumulative_culls: int
    mortality_pct: float
    livability_pct: float
    cumulative_feed_kg: float
    latest_body_weight_g: float | None
    latest_body_weight_date: date | None
    standard_weight_g: float | None
    weight_variance_pct: float | None
    adg_g: float | None  # average daily gain (latest interval)
    cumulative_adg_g: float | None  # average since day 1
    fcr: float | None
    target_min_g: int
    target_max_g: int
    projected_harvest_date: date | None
    projected_harvest_age: int | None


def _latest_weight(records: Iterable[DailyRecord]) -> tuple[date, float] | None:
    weighted = [r for r in records if r.avg_body_weight_g and r.avg_body_weight_g > 0]
    if not weighted:
        return None
    last = max(weighted, key=lambda r: r.record_date)
    return last.record_date, float(last.avg_body_weight_g)


def _project_harvest(
    intake_date: date,
    target_min_g: int,
    target_max_g: int,
    records: list[DailyRecord],
    today: date,
) -> tuple[date | None, int | None]:
    """Project harvest date using recent ADG trend, falling back to standard curve."""
    target_mid = (target_min_g + target_max_g) / 2.0
    weighted = sorted(
        [r for r in records if r.avg_body_weight_g and r.avg_body_weight_g > 0],
        key=lambda r: r.record_date,
    )

    if weighted:
        last = weighted[-1]
        last_age = (last.record_date - intake_date).days
        last_weight = float(last.avg_body_weight_g)

        if last_weight >= target_min_g:
            return last.record_date, last_age

        adg = None
        if len(weighted) >= 2:
            prev = weighted[-2]
            span = (last.record_date - prev.record_date).days
            if span > 0:
                adg = (last_weight - float(prev.avg_body_weight_g)) / span
        if adg is None or adg <= 0:
            std_now = standard_weight_for_age(last_age) or last_weight
            std_next = standard_weight_for_age(last_age + 1)
            if std_next and std_next > std_now:
                adg = std_next - std_now
            else:
                adg = 70.0  # conservative late-stage broiler ADG

        days_needed = max(1, int((target_mid - last_weight) / adg))
        projected = last.record_date + timedelta(days=days_needed)
        return projected, (projected - intake_date).days

    # No weight data — use standard curve to estimate when target_mid is reached.
    for age in range(0, max(COBB500_BODY_WEIGHT_G) + 1):
        if (standard_weight_for_age(age) or 0) >= target_mid:
            return intake_date + timedelta(days=age), age
    return None, None


def compute_batch_metrics(batch: Batch, today: date | None = None) -> BatchMetrics:
    today = today or date.today()
    records = list(batch.daily_records)
    age_days = batch.age_days_on(today)

    cumulative_mortality = sum(r.mortality_count for r in records)
    cumulative_culls = sum(r.cull_count for r in records)
    harvested = sum(h.birds_harvested for h in batch.harvest_events)
    losses = cumulative_mortality + cumulative_culls + harvested
    birds_alive = max(0, batch.intake_count - losses)

    mortality_pct = (
        100.0 * (cumulative_mortality + cumulative_culls) / batch.intake_count
        if batch.intake_count
        else 0.0
    )
    livability_pct = 100.0 - mortality_pct

    cumulative_feed_kg = sum(r.feed_consumed_kg for r in records)

    latest = _latest_weight(records)
    latest_weight_date = latest[0] if latest else None
    latest_weight_g = latest[1] if latest else None

    standard_g = (
        standard_weight_for_age((latest_weight_date - batch.intake_date).days)
        if latest_weight_date
        else None
    )
    variance_pct = (
        100.0 * (latest_weight_g - standard_g) / standard_g
        if (standard_g and latest_weight_g)
        else None
    )

    weighted_records = sorted(
        [r for r in records if r.avg_body_weight_g and r.avg_body_weight_g > 0],
        key=lambda r: r.record_date,
    )
    adg = None
    if len(weighted_records) >= 2:
        a, b = weighted_records[-2], weighted_records[-1]
        span = (b.record_date - a.record_date).days
        if span > 0:
            adg = (b.avg_body_weight_g - a.avg_body_weight_g) / span
    cumulative_adg = None
    if weighted_records:
        last = weighted_records[-1]
        days = (last.record_date - batch.intake_date).days
        if days > 0:
            chick_weight = standard_weight_for_age(0) or 42.0
            cumulative_adg = (last.avg_body_weight_g - chick_weight) / days

    # FCR = total feed consumed / total live weight gain produced (kg).
    fcr = None
    if cumulative_feed_kg > 0 and latest_weight_g and birds_alive > 0:
        chick_weight = standard_weight_for_age(0) or 42.0
        live_weight_kg = (
            birds_alive * (latest_weight_g - chick_weight) / 1000.0
        )
        if live_weight_kg > 0:
            fcr = cumulative_feed_kg / live_weight_kg

    projected_date, projected_age = _project_harvest(
        batch.intake_date,
        batch.target_weight_min_g,
        batch.target_weight_max_g,
        records,
        today,
    )

    return BatchMetrics(
        batch_id=batch.id,
        batch_code=batch.batch_code,
        intake_date=batch.intake_date,
        intake_count=batch.intake_count,
        age_days=age_days,
        birds_alive=birds_alive,
        cumulative_mortality=cumulative_mortality,
        cumulative_culls=cumulative_culls,
        mortality_pct=mortality_pct,
        livability_pct=livability_pct,
        cumulative_feed_kg=cumulative_feed_kg,
        latest_body_weight_g=latest_weight_g,
        latest_body_weight_date=latest_weight_date,
        standard_weight_g=standard_g,
        weight_variance_pct=variance_pct,
        adg_g=adg,
        cumulative_adg_g=cumulative_adg,
        fcr=fcr,
        target_min_g=batch.target_weight_min_g,
        target_max_g=batch.target_weight_max_g,
        projected_harvest_date=projected_date,
        projected_harvest_age=projected_age,
    )


DEFAULT_VACCINE_SCHEDULE = [
    {"age": 7, "name": "Newcastle + IB (LaSota)", "route": "drinking water"},
    {"age": 14, "name": "Gumboro / IBD", "route": "drinking water"},
    {"age": 21, "name": "Newcastle booster", "route": "drinking water"},
]
