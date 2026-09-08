# Rittenregistratie

Deze tool draait onder:

`https://artsjeroen.ddns.net/tools/rittenregistratie/`

## Ontwikkelstappen

1. Frontend-prototype: rit invoeren, km berekenen en CSV-export.
2. GPS -> adres: locatie op telefoon ophalen en via OpenStreetMap Nominatim reverse-geocoden.
3. Raspberry Pi API + SQLite: centrale en duurzame opslag van ritten.
4. Routecontrole: optionele routeafstand vergelijken met kilometertellerafstand.
5. Jaaroverzicht + definitieve export.
6. Voertuiggegevens.
7. Gecontroleerde correctie/audit.
8. Koppeling vanaf de startpagina.

## Architectuur

- GitHub `main` is de bron van waarheid voor code.
- De website staat op de RPi onder `/var/www/html`.
- De ritten-API draait lokaal op `127.0.0.1:8765` via systemd.
- Nginx publiceert de API onder `/tools/rittenregistratie/api/`.
- De SQLite-database staat buiten de Git-repository in `/var/lib/rittenregistratie/ritten.db`.
- Ritdata en voertuigdata horen niet in GitHub.

## Opslaggedrag

De frontend gebruikt geen lokale browseropslag voor ritten of voertuiggegevens. Ritten worden via `GET/POST ./api/rides` gelezen en opgeslagen. De voertuiggegevens worden via `GET/PUT ./api/vehicle` centraal beheerd.

De API controleert dat de beginstand van een nieuwe rit aansluit op de vorige eindstand. Nieuwe ritten kunnen pas worden opgeslagen nadat voertuiggegevens zijn vastgelegd. Verwijderen is voorlopig bewust niet beschikbaar; er komt een gecontroleerde correctie- en auditfunctie zodat wijzigingen aan de registratie traceerbaar blijven.

## Voertuiggegevens

De huidige versie ondersteunt één actief voertuigprofiel met:

- merk;
- type/model;
- kenteken;
- datum in gebruik vanaf;
- optionele datum in gebruik tot.

Deze gegevens worden centraal in SQLite opgeslagen en automatisch bovenaan de jaar-CSV opgenomen. Bij een toekomstige voertuigwissel moet vóór definitief gebruik ondersteuning voor historische voertuigprofielen worden toegevoegd, zodat oude ritten aan het juiste voertuig gekoppeld blijven.

## GPS en adressen

De frontend wordt via HTTPS aangeboden zodat telefoongeolocatie kan worden gebruikt. Reverse geocoding gebeurt alleen wanneer de gebruiker bewust op de locatieknop drukt. Voor de huidige versie wordt de publieke OpenStreetMap Nominatim-service gebruikt; geen automatische of bulk-aanvragen. Kaart- en adresgegevens worden toegeschreven aan OpenStreetMap-bijdragers.

## Routecontrole

De knop `Controleer route` gebruikt de GPS-coördinaten van vertrek en aankomst en vraagt via de eigen backend een normale autoroute op bij OSRM. De controle is adviserend en blokkeert het opslaan van een rit niet als de routingdienst niet bereikbaar is.

Een verschil geldt als opvallend wanneer het groter is dan 3 km of 20% van de berekende routeafstand, waarbij de grootste grens wordt gebruikt. De kilometerteller blijft leidend. Bij een afwijking kan de gebruiker een toelichting of afwijkende route noteren.

## Jaaroverzicht en export

De gebruiker kan per kalenderjaar filteren. De vier samenvattingskaarten en de rittenlijst tonen alleen het gekozen jaar. De knop `Download jaar-CSV` exporteert het gekozen jaar met voertuiggegevens, aantallen, begin- en eindstand, zakelijke kilometers, privékilometers, totaal gereden kilometers en de volledige rittenlijst.

## Huidige status

Stap 1 t/m 6 zijn operationeel. De volgende stap is een gecontroleerde correctie-/auditfunctie. Pas daarna is de registratie functioneel compleet genoeg om als definitieve administratie te gebruiken.
