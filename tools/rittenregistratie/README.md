# Rittenregistratie

Deze tool draait onder:

`https://artsjeroen.ddns.net/tools/rittenregistratie/`

## Navigatie

De interface is opgebouwd als compacte webapp met een hamburgermenu linksboven en vier schermen:

- **Dashboard** — beknopte totalen van het huidige kalenderjaar, actief kenteken en een nieuwe rit toevoegen.
- **Voertuig toevoegen** — nieuw kenteken registreren, RDW-gegevens ophalen en begin-kilometerstand bij ingebruikname vastleggen.
- **Ritcorrectie & auditlog** — bestaande rit gecontroleerd corrigeren en eerdere correcties bekijken.
- **Overzicht** — ritten filteren per kalenderjaar of per maand binnen een jaar en de huidige selectie als CSV exporteren.

## Architectuur

- GitHub `main` is de bron van waarheid voor code.
- De website staat op de RPi onder `/var/www/html`.
- De ritten-API draait lokaal op `127.0.0.1:8765` via systemd.
- Nginx publiceert de API onder `/tools/rittenregistratie/api/`.
- De SQLite-database staat buiten Git in `/var/lib/rittenregistratie/ritten.db`.
- Ritdata, back-ups, mailwachtwoorden en wachtwoordbestanden horen niet in GitHub.

## Voertuigen en kilometerketens

Iedere rit bevat een `vehicleId`. De kilometerketen wordt daardoor per voertuig gecontroleerd en niet over verschillende kentekens heen.

Bij een nieuw voertuig worden merk en handelsbenaming via RDW Open Data opgehaald op basis van het kenteken. De gebruiker legt daarnaast verplicht vast:

- datum `In gebruik vanaf`;
- kilometerstand op die datum;
- optioneel `In gebruik tot`.

De kilometerstand bij ingebruikname vormt het beginpunt van de eigen keten van dat voertuig. De eerste rit moet op die stand beginnen. Daarna moet iedere volgende rit aansluiten op de vorige eindstand van hetzelfde voertuig.

API:

- `GET ./api/vehicles` — alle voertuigen;
- `POST ./api/vehicles` — extra voertuig toevoegen;
- `PUT ./api/vehicles/<id>` — voertuiggegevens wijzigen;
- `GET ./api/rides` — ritten inclusief gekoppeld voertuig/kenteken en beschikbare vertrek-/aankomsttijd;
- `POST ./api/rides` — rit toevoegen;
- `PATCH ./api/rides/<id>` — gecontroleerde correctie;
- `GET ./api/audit` — auditlog.

## Dashboard

Het dashboard toont voor het huidige kalenderjaar:

- aantal ritten;
- zakelijke kilometers;
- privékilometers.

Daarnaast kies je hier het actieve kenteken en voeg je een nieuwe rit toe. De laatste kilometerstand volgt altijd het gekozen voertuig.

## Tijdregistratie

Bij gebruik van de knop **Gebruik locatie** wordt naast de GPS-locatie ook het lokale tijdstip van de telefoon/browser vastgelegd.

- bij het vertrekadres wordt `departureTime` opgeslagen;
- bij het aankomstadres wordt `arrivalTime` opgeslagen;
- de tijd wordt vastgelegd op het moment dat de locatie-opvraag wordt gestart;
- bij een nieuwe locatie-opvraag wordt het tijdstip vervangen door het nieuwe tijdstip;
- als locatiebepaling mislukt, wordt het bij die poging vastgelegde tijdstip niet gebruikt.

De velden worden als `HH:MM:SS` in SQLite opgeslagen. Bestaande ritten van vóór deze wijziging houden lege tijdvelden. Het overzicht toont de tijd bij het vertrek- en aankomstadres en de CSV-export bevat aparte kolommen `Vertrektijd` en `Aankomsttijd`.

## Overzicht en export

Het scherm **Overzicht** heeft twee filters:

- jaar;
- maand, waarbij `Alle maanden` het volledige jaar toont.

Tabel en samenvatting volgen dezelfde selectie. De CSV-export exporteert eveneens precies de gekozen periode. Bij een maandselectie krijgt het bestand bijvoorbeeld de naam `rittenregistratie-2027-03.csv`.

## GPS, adressen en routecontrole

De frontend wordt via HTTPS aangeboden zodat telefoongeolocatie kan worden gebruikt. Reverse geocoding gebeurt alleen wanneer de gebruiker bewust op de locatieknop drukt. De huidige versie gebruikt OpenStreetMap Nominatim.

Na het vastleggen van vertrek en aankomst wordt de routecontrole automatisch uitgevoerd zodra ook de begin- en eindkilometerstand bekend zijn. De controle is adviserend; de kilometerteller blijft leidend.

## Correcties en auditlog

Bestaande ritten worden niet hard verwijderd. Een correctie vereist een reden van minimaal 5 tekens. SQLite bewaart in `ride_audit` de oude en nieuwe versie, reden en correctietijdstip.

Vertrek- en aankomsttijd blijven bij een correctie behouden en maken deel uit van de snapshots in het auditlog. De kilometerketen wordt bij correcties alleen gecontroleerd binnen hetzelfde voertuig/kenteken.

## Dagelijkse back-up

`backup.py` maakt met de SQLite backup-API een consistente kopie van `/var/lib/rittenregistratie/ritten.db` en controleert die met `PRAGMA integrity_check`.

Back-ups staan in `/var/backups/rittenregistratie/`. Standaard worden back-ups ouder dan 35 dagen verwijderd. Systemd gebruikt:

- `deploy/rittenregistratie-backup.service`
- `deploy/rittenregistratie-backup.timer`

## Maandelijkse e-mailrapportage

`monthly_report.py` leest de SQLite-database alleen-lezen en maakt op de eerste dag van iedere maand een rapport van de volledige vorige kalendermaand.

De e-mail bevat:

- aantal ritten;
- zakelijke kilometers;
- privékilometers;
- totaal aantal kilometers;
- een CSV-bijlage met alle ritten, inclusief vertrek-/aankomsttijd en kenteken.

Ook een maand zonder ritten wordt verzonden, met een CSV die alleen de kolomkoppen bevat.

Systemd gebruikt:

- `deploy/rittenregistratie-monthly-report.service`
- `deploy/rittenregistratie-monthly-report.timer`

De timer draait op de eerste dag van de maand rond 08:00 lokale RPi-tijd en gebruikt `Persistent=true`, zodat een gemiste run na een latere boot alsnog wordt uitgevoerd.

Mailconfiguratie staat uitsluitend op de RPi in `/etc/rittenregistratie-mail.env`. Gebruik `deploy/rittenregistratie-mail.env.example` als voorbeeld. Een Gmail app-wachtwoord mag nooit in GitHub of chat worden gezet.

## Releasenotes

Rittenregistratie heeft een eigen leesbare releasenotepagina onder `release/rittenregistratie.html`, gescheiden van de startpagina en andere tools.

## Toegangsbeveiliging

De volledige map `/tools/rittenregistratie/` en de API worden in Nginx met HTTP Basic Auth beveiligd. Het wachtwoordbestand staat alleen op de Raspberry Pi in `/etc/nginx/rittenregistratie.htpasswd`.
