#!/usr/bin/env python3
import csv
import io
import os
import smtplib
import sqlite3
from datetime import date, timedelta
from email.message import EmailMessage

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


def period_bounds(year, month):
    start = date(year, month, 1)
    if month == 12:
        end = date(year + 1, 1, 1)
    else:
        end = date(year, month + 1, 1)
    return start.isoformat(), end.isoformat()


def load_rides(year, month):
    start, end = period_bounds(year, month)
    db = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    db.row_factory = sqlite3.Row
    try:
        return db.execute(
            """
            SELECT
                r.ride_date,
                v.plate AS vehicle_plate,
                r.ride_type,
                r.departure_time,
                r.departure_address,
                r.arrival_time,
                r.arrival_address,
                r.start_odometer,
                r.end_odometer,
                r.distance,
                r.notes
            FROM rides r
            LEFT JOIN vehicles v ON v.id = r.vehicle_id
            WHERE r.ride_date >= ? AND r.ride_date < ?
            ORDER BY r.ride_date ASC, r.id ASC
            """,
            (start, end),
        ).fetchall()
    finally:
        db.close()


def totals(rows):
    business = sum(row["distance"] for row in rows if row["ride_type"] == "business")
    private = sum(row["distance"] for row in rows if row["ride_type"] == "private")
    return business, private


def build_csv(rows):
    output = io.StringIO(newline="")
    writer = csv.writer(output, delimiter=";", quoting=csv.QUOTE_MINIMAL)
    writer.writerow([
        "Datum", "Kenteken", "Type", "Vertrektijd", "Vertrekadres",
        "Aankomsttijd", "Aankomstadres", "Begin km-stand", "Eind km-stand",
        "Kilometers", "Toelichting",
    ])
    for row in rows:
        writer.writerow([
            row["ride_date"],
            row["vehicle_plate"] or "",
            "Privé" if row["ride_type"] == "private" else "Zakelijk",
            row["departure_time"] or "",
            row["departure_address"],
            row["arrival_time"] or "",
            row["arrival_address"],
            row["start_odometer"],
            row["end_odometer"],
            row["distance"],
            row["notes"] or "",
        ])
    return "\ufeff" + output.getvalue()


def send_report(year, month, rows):
    if not SMTP_USER or not SMTP_PASSWORD or not MAIL_TO or not MAIL_FROM:
        raise RuntimeError(
            "Mailconfiguratie ontbreekt: stel RITTEN_SMTP_USER, RITTEN_SMTP_PASSWORD, "
            "RITTEN_MAIL_TO en eventueel RITTEN_MAIL_FROM in"
        )

    business_km, private_km = totals(rows)
    total_km = business_km + private_km
    month_name = MONTHS_NL[month - 1]
    subject = f"Rittenregistratie {month_name} {year}"
    filename = f"rittenregistratie-{year}-{month:02d}.csv"

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = MAIL_FROM
    message["To"] = MAIL_TO
    message.set_content(
        f"Maandoverzicht rittenregistratie – {month_name} {year}\n\n"
        f"Aantal ritten: {len(rows)}\n"
        f"Zakelijke kilometers: {business_km} km\n"
        f"Privékilometers: {private_km} km\n"
        f"Totaal: {total_km} km\n\n"
        f"De volledige rittenlijst staat als CSV in de bijlage.\n"
    )
    csv_bytes = build_csv(rows).encode("utf-8")
    message.add_attachment(
        csv_bytes,
        maintype="text",
        subtype="csv",
        filename=filename,
    )

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
        smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(message)

    print(
        f"Maandrapport verzonden: {month_name} {year}, {len(rows)} ritten, "
        f"{business_km} zakelijke km, {private_km} privé km -> {MAIL_TO}"
    )


def main():
    year, month = previous_month()
    rows = load_rides(year, month)
    send_report(year, month, rows)


if __name__ == "__main__":
    main()
