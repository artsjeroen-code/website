#!/usr/bin/env python3
import os
import re
import smtplib
import sqlite3
from datetime import date, timedelta
from email.message import EmailMessage

from employer_report import build_workbook, cumulative_totals

DB_PATH = os.environ.get("RITTEN_DB", "/var/lib/rittenregistratie/ritten.db")
SMTP_HOST = os.environ.get("RITTEN_SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("RITTEN_SMTP_PORT", "465"))
SMTP_USER = os.environ.get("RITTEN_SMTP_USER", "").strip()
SMTP_PASSWORD = os.environ.get("RITTEN_SMTP_PASSWORD", "").replace(" ", "")
MAIL_TO = os.environ.get("RITTEN_MAIL_TO", SMTP_USER).strip()
MAIL_FROM = os.environ.get("RITTEN_MAIL_FROM", SMTP_USER).strip()

MONTHS_NL = [
    "januari", "februari", "maart", "april", "mei", "juni",
    "juli", "augustus", "september", "oktober", "november", "december",
]


def previous_month(today=None):
    today = today or date.today()
    first_this_month = today.replace(day=1)
    last_previous_month = first_this_month - timedelta(days=1)
    return last_previous_month.year, last_previous_month.month


def first_day_after(year, month):
    if month == 12:
        return date(year + 1, 1, 1).isoformat()
    return date(year, month + 1, 1).isoformat()


def load_data(cutoff):
    db = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    db.row_factory = sqlite3.Row
    try:
        vehicles = [
            {
                "id": row["id"],
                "make": row["make"],
                "model": row["model"],
                "plate": row["plate"],
                "useFrom": row["use_from"],
                "useTo": row["use_to"],
                "initialOdometer": row["initial_odometer"],
            }
            for row in db.execute(
                "SELECT * FROM vehicles ORDER BY use_from ASC, id ASC"
            ).fetchall()
        ]
        rides = [
            {
                "id": row["id"],
                "vehicleId": row["vehicle_id"],
                "date": row["ride_date"],
                "type": row["ride_type"],
                "startOdometer": row["start_odometer"],
                "endOdometer": row["end_odometer"],
                "distance": row["distance"],
                "departureTime": row["departure_time"],
                "departureAddress": row["departure_address"],
                "arrivalTime": row["arrival_time"],
                "arrivalAddress": row["arrival_address"],
                "notes": row["notes"],
            }
            for row in db.execute(
                """
                SELECT id, vehicle_id, ride_date, ride_type, start_odometer, end_odometer,
                       distance, departure_time, departure_address, arrival_time,
                       arrival_address, notes
                FROM rides
                WHERE ride_date < ?
                ORDER BY ride_date ASC, id ASC
                """,
                (cutoff,),
            ).fetchall()
        ]
        return vehicles, rides
    finally:
        db.close()


def vehicles_for_year(vehicles, year):
    start = f"{year}-01-01"
    end = f"{year + 1}-01-01"
    return [
        vehicle for vehicle in vehicles
        if str(vehicle.get("useFrom") or "") < end
        and (not vehicle.get("useTo") or str(vehicle["useTo"]) >= start)
    ]


def safe_plate(value):
    return re.sub(r"[^A-Za-z0-9-]", "", str(value or "voertuig")) or "voertuig"


def send_report(year, through_month, vehicles, rides):
    if not SMTP_USER or not SMTP_PASSWORD or not MAIL_TO or not MAIL_FROM:
        raise RuntimeError(
            "Mailconfiguratie ontbreekt: stel RITTEN_SMTP_USER, RITTEN_SMTP_PASSWORD, "
            "RITTEN_MAIL_TO en eventueel RITTEN_MAIL_FROM in"
        )

    report_vehicles = vehicles_for_year(vehicles, year)
    month_name = MONTHS_NL[through_month - 1]
    subject = f"Rittenregistratie {year} t/m {month_name}"
    total_rides = 0
    total_business = 0
    total_private = 0
    attachments = []

    for vehicle in report_vehicles:
        vehicle_rides = [
            ride for ride in rides
            if int(ride.get("vehicleId") or 0) == int(vehicle["id"])
        ]
        ride_count, business_km, private_km = cumulative_totals(
            vehicle, vehicle_rides, year, through_month
        )
        total_rides += ride_count
        total_business += business_km
        total_private += private_km
        attachments.append((
            f"Rittenregistratie-{year}-{safe_plate(vehicle.get('plate'))}.xlsx",
            build_workbook(vehicle, vehicle_rides, year),
        ))

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = MAIL_FROM
    message["To"] = MAIL_TO
    message.set_content(
        f"Cumulatieve rittenregistratie {year} t/m {month_name}.\n\n"
        f"Aantal ritten: {total_rides}\n"
        f"Zakelijke kilometers: "
        f"{int(total_business) if float(total_business).is_integer() else total_business} km\n"
        f"Privékilometers: "
        f"{int(total_private) if float(total_private).is_integer() else total_private} km\n\n"
        f"De bijlage(n) volgen het werkgeversformat en bevatten twaalf maandtabbladen. "
        f"Alle geregistreerde gegevens tot en met {month_name} zijn opgenomen; "
        f"latere maanden blijven leeg.\n"
    )

    for filename, workbook in attachments:
        message.add_attachment(
            workbook,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
        )

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
        smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(message)

    print(
        f"Cumulatief werkgeversrapport verzonden: {year} t/m {month_name}, "
        f"{total_rides} ritten, {len(attachments)} Excel-bijlage(n) -> {MAIL_TO}"
    )


def main():
    year, through_month = previous_month()
    cutoff = first_day_after(year, through_month)
    vehicles, rides = load_data(cutoff)
    send_report(year, through_month, vehicles, rides)


if __name__ == "__main__":
    main()
