# Rittenregistratie

Deze tool draait onder:

`https://artsjeroen.ddns.net/tools/rittenregistratie/`

## Ontwikkelstappen

1. Frontend-prototype: rit invoeren, km berekenen, lokale testopslag en CSV-export.
2. GPS -> adres: locatie op telefoon ophalen en via OpenStreetMap Nominatim reverse-geocoden.
3. Raspberry Pi API + SQLite: centrale en duurzame opslag van ritten.
4. Frontend koppelen aan de API en lokale testopslag uitfaseren.
5. Routecontrole: optionele routeafstand vergelijken met kilometertellerafstand.
6. Jaaroverzicht + definitieve export.
7. Koppeling vanaf de startpagina.

## Architectuur

- GitHub `main` is de bron van waarheid voor code.
- De statische frontend staat onder `/var/www/html/tools/rittenregistratie/`.
- De Python API draait alleen lokaal op `127.0.0.1:8765`.
- Nginx publiceert de API onder `/tools/rittenregistratie/api/`.
- Productiedata staat buiten Git in `/var/lib/rittenregistratie/ritten.db`.
- De database hoort nooit in GitHub te worden opgenomen.

## API

De eerste API-versie ondersteunt:

- `GET /api/health` — controle of API en database bereikbaar zijn.
- `GET /api/rides` — alle ritten ophalen.
- `POST /api/rides` — rit toevoegen met controle op aansluitende kilometerstand.

De frontend gebruikt deze API nog niet; eerst wordt de API afzonderlijk op de Raspberry Pi getest.

## Deploymentbestanden

- `server.py` — Python/SQLite API zonder externe Python-pakketten.
- `deploy/rittenregistratie.service` — systemd-service voor de API.
- `deploy/nginx-location.conf` — Nginx reverse-proxyfragment.

## Belangrijk

- De ritgegevens zelf horen niet in GitHub.
- Maak later periodiek een back-up van `/var/lib/rittenregistratie/ritten.db`.
- De frontend wordt via HTTPS aangeboden zodat telefoongeolocatie kan worden gebruikt.
- Reverse geocoding gebeurt alleen wanneer de gebruiker bewust op de locatieknop drukt.
- Voor de huidige testfase wordt de publieke OpenStreetMap Nominatim-service gebruikt. Dit gebruik moet beperkt blijven; geen automatische of bulk-aanvragen.
- Kaart- en adresgegevens worden toegeschreven aan OpenStreetMap-bijdragers.

## Huidige status

Stap 1 en 2 werken als frontendprototype. De API- en SQLite-code voor stap 3 staat in GitHub `main` en moet op de Raspberry Pi worden geïnstalleerd en getest. Totdat stap 4 is afgerond, bewaart de frontend ritten nog in lokale browseropslag.
