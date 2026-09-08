# Rittenregistratie

Deze tool draait onder:

`https://artsjeroen.ddns.net/tools/rittenregistratie/`

## Ontwikkelstappen

1. Frontend-prototype: rit invoeren, km berekenen en CSV-export.
2. GPS -> adres: locatie op telefoon ophalen en via OpenStreetMap Nominatim reverse-geocoden.
3. Raspberry Pi API + SQLite: centrale en duurzame opslag van ritten.
4. Routecontrole: optionele routeafstand vergelijken met kilometertellerafstand.
5. Jaaroverzicht + definitieve export.
6. Koppeling vanaf de startpagina.

## Architectuur

- GitHub `main` is de bron van waarheid voor code.
- De website staat op de RPi onder `/var/www/html`.
- De ritten-API draait lokaal op `127.0.0.1:8765` via systemd.
- Nginx publiceert de API onder `/tools/rittenregistratie/api/`.
- De SQLite-database staat buiten de Git-repository in `/var/lib/rittenregistratie/ritten.db`.
- Ritdata hoort niet in GitHub.

## Opslaggedrag

De frontend gebruikt geen lokale browseropslag meer voor ritten. Bij laden worden ritten via `GET ./api/rides` opgehaald en nieuwe ritten worden via `POST ./api/rides` centraal opgeslagen.

De API controleert dat de beginstand van een nieuwe rit aansluit op de vorige eindstand. Verwijderen is voorlopig bewust niet beschikbaar; er komt later een gecontroleerde correctie- en auditfunctie zodat wijzigingen aan een fiscale registratie traceerbaar blijven.

## GPS en adressen

De frontend wordt via HTTPS aangeboden zodat telefoongeolocatie kan worden gebruikt. Reverse geocoding gebeurt alleen wanneer de gebruiker bewust op de locatieknop drukt. Voor de huidige versie wordt de publieke OpenStreetMap Nominatim-service gebruikt; geen automatische of bulk-aanvragen. Kaart- en adresgegevens worden toegeschreven aan OpenStreetMap-bijdragers.

## Routecontrole

De knop `Controleer route` gebruikt de GPS-coördinaten van vertrek en aankomst en vraagt via de eigen backend een normale autoroute op bij OSRM. De OSRM Route service retourneert de afstand van de berekende route. De controle is adviserend en blokkeert het opslaan van een rit niet als de routingdienst niet bereikbaar is.

Een verschil geldt als opvallend wanneer het groter is dan 3 km of 20% van de berekende routeafstand, waarbij de grootste grens wordt gebruikt. De kilometerteller blijft leidend. Bij een afwijking kan de gebruiker een toelichting of afwijkende route noteren.

De publieke OSRM-demo wordt alleen gebruikt na een bewuste druk op de routecontroleknop en niet voor bulk- of achtergrondverkeer. Routegegevens worden toegeschreven aan OSRM/OpenStreetMap.

## Huidige status

Stap 1 t/m 4 zijn operationeel. Ritten worden centraal in SQLite op de Raspberry Pi opgeslagen en zijn vanaf verschillende apparaten zichtbaar. De volgende ontwikkelstap is het jaaroverzicht en de definitieve exportfuncties.
