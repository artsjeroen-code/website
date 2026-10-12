#!/usr/bin/env python3
"""Publiceer alleen de wijzigingsdatum van de gehoste werkagenda.

Het geheime pad naar het ICS-bestand komt nooit in het statusbestand,
de loguitvoer of de startpagina terecht.
"""

import argparse
import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ICS_ROOT = Path("/var/www/html")
ALLOWED_OUTPUT_ROOTS = {
    Path("/var/www/staging"),
    Path("/var/www/html"),
}
STATUS_FILENAME = "calendar-status.json"


def main():
    parser = argparse.ArgumentParser(
        description="Maak een openbare status met alleen de ICS-wijzigingsdatum."
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        required=True,
        help="De webroot: /var/www/staging of /var/www/html",
    )
    args = parser.parse_args()

    output_root = args.output_root.resolve()
    if output_root not in ALLOWED_OUTPUT_ROOTS:
        parser.error("Ongeldige output-root.")
    if not output_root.is_dir():
        parser.error("De output-root bestaat niet.")

    # Geen geheime hexmap opnemen in code, status of logs.
    calendars = [
        path for path in ICS_ROOT.glob("*/Werkagenda.ics")
        if path.is_file()
    ]
    if len(calendars) != 1:
        raise SystemExit(
            "Verwacht precies één gehost Werkagenda.ics-bestand; "
            f"gevonden: {len(calendars)}. Status niet bijgewerkt."
        )

    modified = datetime.fromtimestamp(
        calendars[0].stat().st_mtime, tz=timezone.utc
    ).isoformat()
    payload = {"last_modified": modified}

    # Publiceer atomair, zodat de browser nooit gedeeltelijke JSON ziet.
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            prefix=".calendar-status-",
            suffix=".tmp",
            dir=output_root,
            delete=False,
        ) as output:
            temporary_path = Path(output.name)
            json.dump(payload, output, ensure_ascii=False)
            output.write("\n")
            output.flush()
            os.fsync(output.fileno())

        os.chmod(temporary_path, 0o644)
        os.replace(temporary_path, output_root / STATUS_FILENAME)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)

    print("Agendastatus bijgewerkt (alleen datum/tijd).", flush=True)


if __name__ == "__main__":
    main()
