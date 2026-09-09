# Rittenregistratie

Deze tool draait onder:

`https://artsjeroen.ddns.net/tools/rittenregistratie/`

## Ontwikkelstappen

1. Frontend-prototype: rit invoeren, km berekenen en CSV-export.
2. GPS -> adres: locatie op telefoon ophalen en via OpenStreetMap Nominatim reverse-geocoden.
3. Raspberry Pi API + SQLite: centrale en duurzame opslag van ritten.
4. Routecontrole: optionele routeafstand vergelijken met kilometertellerafstand.
5. Jaaroverzicht + definitieve export.
6. Meerdere voertuigen en kilometerketens per kenteken.
7. RDW-lookup + beginstand per voertuig.
8. Gecontroleerde correcties + auditlog.
9. Dagelijkse SQLite-back-up + hersteltest.
10. Toegangsbeveiliging via Nginx Basic Auth.
11. Koppeling vanaf de startpagina.

## Architectuur

- GitHub `main` is de bron van waarheid voor code.
- De website staat op de RPi onder `/var/www/html`.
- De ritten-API draait lokaal op `127.0.0.1:8765` via systemd.
- Nginx publiceert de API onder `/tools/rittenregistratie/api/`.
- De SQLite-database staat buiten de Git-repository in `/var/lib/rittenregistratie/ritten.db`.
- Ritdata, back-ups en wachtwoordbestanden horen niet in GitHub.

## Opslaggedrag

De frontend gebruikt geen lokale browseropslag voor ritten. Bij laden worden ritten via `GET ./api/rides` opgehaald en nieuwe ritten via `POST ./api/rides` centraal opgeslagen.

Iedere rit bevat een `vehicleId`. De kilometerketen wordt daardoor per voertuig gecontroleerd en niet meer over alle ritten heen.

## Voertuigen en kentekens

De interface toont standaard alleen een compacte kentekenkeuze. Het volledige formulier verschijnt pas na `Voertuig toevoegen`.

Bij een nieuw voertuig wordt eerst het kenteken ingevoerd. De frontend vraagt vervolgens de officiële RDW Open Data-dataset `Gekentekende_voertuigen` (`m9d7-ebf2`) op en vult `merk` en `handelsbenaming` automatisch in. De RDW levert geen persoonlijke tellerstand voor deze administratie; de gebruiker legt daarom zelf verplicht de kilometerstand vast die hoort bij `In gebruik vanaf`.

Die kilometerstand wordt centraal opgeslagen als `initial_odometer` en vormt het startpunt van de kilometerketen van dat voertuig. De eerste rit van het voertuig moet op die stand beginnen. Daarna moet iedere volgende rit aansluiten op de vorige eindstand van hetzelfde kenteken.

Voertuigen worden opgeslagen in de tabel `vehicles`. De bestaande oude tabel `vehicle` blijft alleen aanwezig voor migratie/compatibiliteit. Bij de eerste start na de meervoertuigenmigratie wordt het bestaande voertuig automatisch naar `vehicles` gemigreerd en worden bestaande ritten daaraan gekoppeld. Voor bestaande voertuigen wordt de beginstand waar mogelijk afgeleid uit de eerste reeds opgeslagen rit.

Een nieuw voertuig begint altijd een eigen kilometerketen. De beginstand van een nieuw kenteken hoeft dus niet aan te sluiten op de eindstand van een ander voertuig.

API:

- `GET ./api/vehicles` — alle voertuigen;
- `POST ./api/vehicles` — extra voertuig toevoegen inclusief beginstand;
- `PUT ./api/vehicles/<id>` — voertuiggegevens wijzigen;
- `GET ./api/vehicle` — tijdelijke compatibiliteitsroute die het laatst toegevoegde voertuig teruggeeft.

## GPS en adressen

De frontend wordt via HTTPS aangeboden zodat telefoongeolocatie kan worden gebruikt. Reverse geocoding gebeurt alleen wanneer de gebruiker bewust op de locatieknop drukt. Voor de huidige versie wordt de publieke OpenStreetMap Nominatim-service gebruikt; geen automatische of bulk-aanvragen. Kaart- en adresgegevens worden toegeschreven aan OpenStreetMap-bijdragers.

## Routecontrole

De knop `Controleer route` gebruikt de GPS-coördinaten van vertrek en aankomst en vraagt via de eigen backend een normale autoroute op bij OSRM. De controle is adviserend en blokkeert het opslaan van een rit niet als de routingdienst niet bereikbaar is. De kilometerteller blijft leidend.

## Jaaroverzicht en export

De gebruiker kan per kalenderjaar filteren. De samenvattingskaarten, rittenlijst en CSV-export volgen het gekozen jaar. De rittenlijst bevat ook het kenteken. De jaar-CSV bevat de voertuigen die in dat jaar voorkomen, inclusief beginstand bij ingebruikname, plus per rit het kenteken, type, adressen en kilometerstanden.

Zakelijke en privékilometers worden over alle voertuigen in het gekozen jaar opgeteld. De kaart `Laatste km-stand gekozen auto` volgt het voertuig dat bovenaan geselecteerd is.

## Correcties en auditlog

Bestaande ritten worden niet hard verwijderd. Een fout wordt gecorrigeerd via `PATCH ./api/rides/<id>` en vereist altijd een reden van minimaal 5 tekens.

Bij iedere correctie bewaart SQLite in `ride_audit` het ritnummer, tijdstip, reden en de volledige oude en nieuwe versie van de rit.

Voor een correctie wordt alleen de kilometerketen van hetzelfde voertuig opnieuw gecontroleerd. Bij de eerste rit van een voertuig wordt ook de vastgelegde beginstand bij ingebruikname gecontroleerd.

De auditlog is leesbaar via `GET ./api/audit` en wordt in de webinterface getoond.

## Dagelijkse back-up

`backup.py` maakt met de SQLite backup-API een consistente kopie van `/var/lib/rittenregistratie/ritten.db`. De tijdelijke kopie wordt met `PRAGMA integrity_check` gecontroleerd voordat deze als geldige back-up wordt gepubliceerd.

Back-ups staan buiten Git in `/var/backups/rittenregistratie/`. Standaard worden back-ups ouder dan 35 dagen verwijderd. Dit is instelbaar met `RITTEN_BACKUP_RETENTION_DAYS`.

Systemd gebruikt:

- `deploy/rittenregistratie-backup.service`
- `deploy/rittenregistratie-backup.timer`

De timer plant dagelijks rond 03:15 lokale systeemtijd met maximaal 10 minuten willekeurige vertraging. `Persistent=true` zorgt dat een gemiste uitvoering na een uitgeschakelde RPi bij de volgende start alsnog wordt ingehaald.

## Toegangsbeveiliging

De volledige map `/tools/rittenregistratie/` en de specifiekere API-location `/tools/rittenregistratie/api/` worden in Nginx met HTTP Basic Auth beveiligd. Omdat de site uitsluitend via HTTPS wordt gebruikt, worden de Basic Auth-gegevens versleuteld over TLS verzonden.

Het wachtwoordbestand staat alleen op de Raspberry Pi in `/etc/nginx/rittenregistratie.htpasswd` en wordt niet in GitHub opgeslagen. De voorbeeldconfig staat in `deploy/nginx-location.conf`.

## Huidige status

De kernregistratie ondersteunt meerdere voertuigen met een eigen kilometerketen per kenteken. Nieuwe voertuigen kunnen via RDW Open Data automatisch met merk/type worden aangevuld en krijgen een expliciete tellerstand bij ingebruikname als startpunt van hun keten. Centrale opslag, GPS/adressen, routecontrole, jaaroverzicht/export, auditlog, dagelijkse back-up en toegangsbeveiliging blijven behouden.
