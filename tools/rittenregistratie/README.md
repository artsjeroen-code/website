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
- De passkey-authenticatieservice draait lokaal op `127.0.0.1:8766` via systemd.
- Nginx publiceert de tool en gebruikt een interne `auth_request` naar de authenticatieservice.
- De ritten-SQLite-database staat buiten Git in `/var/lib/rittenregistratie/ritten.db`.
- Passkeys en sessies staan apart buiten Git in `/var/lib/rittenregistratie/auth.db`.
- Ritdata, passkeycredentials, back-ups, mailwachtwoorden en wachtwoordbestanden horen niet in GitHub.

## PWA

De rittenregistratie is een online-first Progressive Web App.

Bestanden:

- `manifest.webmanifest` — appnaam, standalone weergave, kleuren en iconen;
- `service-worker.js` — cachet alleen de statische app-shell;
- `pwa.js` — registreert de service worker;
- `icons/` — app-iconen.

Belangrijk: `/tools/rittenregistratie/api/` wordt expliciet niet door de service worker onderschept of gecachet. Ritdata, voertuigen, auditgegevens en authenticatieresponses blijven daardoor uitsluitend via de centrale services lopen. Loginredirects worden niet als app-shell gecachet.

Als de verbinding wegvalt kan een eerder geladen app-shell nog openen, maar gegevens ophalen, ritten opslaan, inloggen, RDW, reverse geocoding en routecontrole vereisen een werkende netwerkverbinding. Er is bewust nog geen offline synchronisatiewachtrij om dubbele ritten of conflicten in de kilometerketen te voorkomen.

## Passkey / Face ID / vingerafdruk

De dagelijkse toegang gebruikt WebAuthn/passkeys. De gebruiker opent eerst `login.html` en authenticatie gebeurt met de passkey-provider van het apparaat. `userVerification=required` wordt gebruikt. Afhankelijk van het platform kan dit Face ID, Touch ID, vingerafdruk, Windows Hello of de apparaat-PIN/-code zijn. De website kan niet afdwingen welke lokale verificatiemethode het besturingssysteem kiest.

De server bewaart geen biometrische informatie en geen private key. In `auth.db` staan alleen de publieke WebAuthn-credentialdata en sessies.

Bestanden:

- `auth_server.py` — aparte WebAuthn- en sessieservice;
- `login.html`, `passkey.js`, `auth.css` — publieke logininterface zonder ritdata;
- `register.html`, `register-passkey.js` — bootstrap/herstel voor nieuwe passkeys;
- `requirements-auth.txt` — gepinde Python-dependency `fido2==1.2.0`;
- `deploy/rittenregistratie-auth.service` — systemd-service op `127.0.0.1:8766`;
- `deploy/nginx-location.conf` — Nginx-routing en sessiecontrole.

De bestaande Nginx Basic Auth blijft alleen voor `register.html` en de registratie-endpoints bestaan. Daardoor kan een eerste of extra passkey alleen worden toegevoegd nadat het bestaande wachtwoord is ingevoerd. Normale toegang tot de rittenregistratie gebruikt daarna geen Basic Auth meer.

Een succesvolle passkey-login maakt een `HttpOnly`, `Secure`, `SameSite=Strict` sessiecookie met scope `/tools/rittenregistratie/`. De standaard systemd-config laat sessies zeven dagen geldig zijn. De normale ritten-API is alleen bereikbaar wanneer Nginx via `/auth-check` een geldige sessie bevestigt.

### Veilige deployvolgorde passkeys

Voer de overstap in deze volgorde uit zodat de huidige Basic Auth actief blijft tot de nieuwe service bewezen werkt:

1. `git pull origin main` op de RPi.
2. Maak een back-up van de rittenregistratie.
3. Installeer een aparte virtualenv en dependency:

   ```bash
   sudo apt update
   sudo apt install python3-venv
   sudo python3 -m venv /var/lib/rittenregistratie/auth-venv
   sudo /var/lib/rittenregistratie/auth-venv/bin/pip install -r /var/www/html/tools/rittenregistratie/requirements-auth.txt
   sudo chown -R www-data:www-data /var/lib/rittenregistratie/auth-venv
   ```

4. Installeer en start eerst alleen de auth-service:

   ```bash
   sudo cp /var/www/html/tools/rittenregistratie/deploy/rittenregistratie-auth.service /etc/systemd/system/
   sudo systemctl daemon-reload
   sudo systemctl enable --now rittenregistratie-auth.service
   systemctl status rittenregistratie-auth.service --no-pager
   curl http://127.0.0.1:8766/api/auth/status
   ```

5. Pas pas na een succesvolle lokale statuscheck de Nginx-locations toe vanuit `deploy/nginx-location.conf`.
6. Voer `sudo nginx -t` uit; alleen bij succes `sudo systemctl reload nginx`.
7. Open `https://artsjeroen.ddns.net/tools/rittenregistratie/`. Zonder sessie moet de browser naar `login.html` gaan.
8. Kies **Eerste passkey instellen**. Alleen dan verschijnt de bestaande Basic Auth-vraag. Registreer de passkey en controleer dat de app daarna opent.
9. Test in een privévenster: normale toegang moet zonder wachtwoordprompt naar de passkey-login gaan; de ritten-API moet zonder sessie `401` geven.

Verwijder het bestaande htpasswd-bestand niet: het blijft het gecontroleerde bootstrap/herstelpad voor het registreren van extra passkeys.

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
