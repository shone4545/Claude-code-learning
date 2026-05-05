from __future__ import annotations

from datetime import date, datetime
from sqlalchemy import (
    Integer,
    String,
    Date,
    DateTime,
    Float,
    Text,
    ForeignKey,
    Index,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .extensions import db


class Batch(db.Model):
    __tablename__ = "batches"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_code: Mapped[str] = mapped_column(String(40), unique=True, index=True)
    intake_date: Mapped[date] = mapped_column(Date, index=True)
    intake_count: Mapped[int] = mapped_column(Integer)
    breed: Mapped[str] = mapped_column(String(50), default="Cobb 500")
    house: Mapped[str] = mapped_column(String(50), default="")
    target_weight_min_g: Mapped[int] = mapped_column(Integer, default=2100)
    target_weight_max_g: Mapped[int] = mapped_column(Integer, default=2500)
    status: Mapped[str] = mapped_column(String(20), default="active", index=True)
    notes: Mapped[str] = mapped_column(Text, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    daily_records: Mapped[list["DailyRecord"]] = relationship(
        back_populates="batch",
        cascade="all, delete-orphan",
        order_by="DailyRecord.record_date",
    )
    vaccine_events: Mapped[list["VaccineEvent"]] = relationship(
        back_populates="batch",
        cascade="all, delete-orphan",
        order_by="VaccineEvent.scheduled_date",
    )
    manpower_logs: Mapped[list["ManpowerLog"]] = relationship(
        back_populates="batch", cascade="all, delete-orphan"
    )
    utility_readings: Mapped[list["UtilityReading"]] = relationship(
        back_populates="batch", cascade="all, delete-orphan"
    )
    harvest_events: Mapped[list["HarvestEvent"]] = relationship(
        back_populates="batch",
        cascade="all, delete-orphan",
        order_by="HarvestEvent.harvest_date",
    )

    def age_days_on(self, on_date: date) -> int:
        return (on_date - self.intake_date).days


class DailyRecord(db.Model):
    __tablename__ = "daily_records"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_id: Mapped[int] = mapped_column(ForeignKey("batches.id"), index=True)
    record_date: Mapped[date] = mapped_column(Date, index=True)
    mortality_count: Mapped[int] = mapped_column(Integer, default=0)
    cull_count: Mapped[int] = mapped_column(Integer, default=0)
    feed_consumed_kg: Mapped[float] = mapped_column(Float, default=0.0)
    water_consumed_l: Mapped[float] = mapped_column(Float, default=0.0)
    avg_body_weight_g: Mapped[float | None] = mapped_column(Float, nullable=True)
    sample_size: Mapped[int | None] = mapped_column(Integer, nullable=True)
    temperature_c: Mapped[float | None] = mapped_column(Float, nullable=True)
    humidity_pct: Mapped[float | None] = mapped_column(Float, nullable=True)
    notes: Mapped[str] = mapped_column(Text, default="")

    batch: Mapped[Batch] = relationship(back_populates="daily_records")

    __table_args__ = (
        Index("ux_daily_batch_date", "batch_id", "record_date", unique=True),
    )


class VaccineEvent(db.Model):
    __tablename__ = "vaccine_events"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_id: Mapped[int] = mapped_column(ForeignKey("batches.id"), index=True)
    scheduled_date: Mapped[date] = mapped_column(Date, index=True)
    age_days: Mapped[int] = mapped_column(Integer)
    vaccine_name: Mapped[str] = mapped_column(String(100))
    route: Mapped[str] = mapped_column(String(40), default="drinking water")
    doses_planned: Mapped[int] = mapped_column(Integer, default=0)
    doses_used: Mapped[int | None] = mapped_column(Integer, nullable=True)
    administered_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    notes: Mapped[str] = mapped_column(Text, default="")

    batch: Mapped[Batch] = relationship(back_populates="vaccine_events")


class ManpowerLog(db.Model):
    __tablename__ = "manpower_logs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    log_date: Mapped[date] = mapped_column(Date, index=True)
    batch_id: Mapped[int | None] = mapped_column(
        ForeignKey("batches.id"), nullable=True, index=True
    )
    worker_name: Mapped[str] = mapped_column(String(80))
    role: Mapped[str] = mapped_column(String(50), default="farmhand")
    hours: Mapped[float] = mapped_column(Float, default=0.0)
    task: Mapped[str] = mapped_column(String(120), default="")
    notes: Mapped[str] = mapped_column(Text, default="")

    batch: Mapped[Batch | None] = relationship(back_populates="manpower_logs")


class UtilityReading(db.Model):
    __tablename__ = "utility_readings"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    reading_date: Mapped[date] = mapped_column(Date, index=True)
    batch_id: Mapped[int | None] = mapped_column(
        ForeignKey("batches.id"), nullable=True, index=True
    )
    utility_type: Mapped[str] = mapped_column(String(20))  # electricity, water, gas, fuel
    quantity: Mapped[float] = mapped_column(Float, default=0.0)
    unit: Mapped[str] = mapped_column(String(10), default="kWh")
    cost: Mapped[float] = mapped_column(Float, default=0.0)
    notes: Mapped[str] = mapped_column(Text, default="")

    batch: Mapped[Batch | None] = relationship(back_populates="utility_readings")


class HarvestEvent(db.Model):
    __tablename__ = "harvest_events"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_id: Mapped[int] = mapped_column(ForeignKey("batches.id"), index=True)
    harvest_date: Mapped[date] = mapped_column(Date, index=True)
    birds_harvested: Mapped[int] = mapped_column(Integer, default=0)
    total_weight_kg: Mapped[float] = mapped_column(Float, default=0.0)
    avg_weight_g: Mapped[float | None] = mapped_column(Float, nullable=True)
    notes: Mapped[str] = mapped_column(Text, default="")

    batch: Mapped[Batch] = relationship(back_populates="harvest_events")
