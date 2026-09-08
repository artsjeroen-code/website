# Rittenregistratie

Deze tool draait onder:

`https://artsjeroen.ddns.net/tools/rittenregistratie/`

## Ontwikkelstappen

1. Frontend-prototype: rit invoeren, km berekenen, lokale testopslag en CSV-export.
2. GPS -> adres: locatie op telefoon ophalen en via OpenStreetMap Nominatim reverse-geocoden.
3. Raspberry Pi API + SQLite: centrale en duurzame opslag van ritten.
4. Routecontrole: optionele routeafstand vergelijken met kilometertellerafstand.
5. Jaaroverzicht + definitieve export.
6. Koppeling vanaf de startpagina.

## Belangrijk

- GitHub `main` is de bron van waarheid voor code.
- De ritgegevens zelf horen niet in GitHub.
- Productiedata wordt later op de Raspberry Pi in SQLite opgeslagen, met back-up.
- De frontend wordt via HTTPS aangeboden zodat telefoongeolocatie kan worden gebruikt.
- Reverse geocoding gebeurt alleen wanneer de gebruiker bewust op de locatieknop drukt.
- Voor de huidige testfase wordt de publieke OpenStreetMap Nominatim-service gebruikt. Dit gebruik moet beperkt blijven; geen automatische of bulk-aanvragen.
- Kaart- en adresgegevens worden toegeschreven aan OpenStreetMap-bijdragers.

## Huidige status

Stap 1 en 2 zijn als prototype beschikbaar. Ritten worden nog alleen in lokale browseropslag bewaard; dit is nadrukkelijk nog niet geschikt als definitieve fiscale administratie. De volgende stap is centrale SQLite-opslag op de Raspberry Pi.
