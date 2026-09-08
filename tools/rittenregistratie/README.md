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
7. Gecontroleerde correcties + auditlog.
8. Dagelijkse SQLite-back-up + hersteltest.
9. Koppeling vanaf de startpagina en beveiliging verder aanscherpen.

## Architectuur

- GitHub `main` is de bron van waarheid voor code.
- De website staat op de RPi onder `/var/www/html`.
- De ritten-API draait lokaal op `127.0.0.1:8765` via systemd.
- Nginx publiceert de API onder `/tools/rittenregistratie/api/`.
- De SQLite-database staat buiten de Git-repository in `/var/lib/rittenregistratie/ritten.db`.
- Ritdata en back-ups horen niet in GitHub.

## Opslaggedrag

De frontend gebruikt geen lokale browseropslag voor ritten. Bij laden worden ritten via `GET ./api/rides` opgehaald en nieuwe ritten via `POST ./api/rides` centraal opgeslagen.

De API controleert dat de beginstand van een nieuwe rit aansluit op de vorige eindstand.

## Voertuiggegevens

Merk, type/model, kenteken en gebruiksperiode worden centraal opgeslagen. Nieuwe ritten kunnen pas worden geregistreerd nadat voertuiggegevens aanwezig zijn. De jaar-CSV neemt deze gegevens mee.

## GPS en adressen

De frontend wordt via HTTPS aangeboden zodat telefoongeolocatie kan worden gebruikt. Reverse geocoding gebeurt alleen wanneer de gebruiker bewust op de locatieknop drukt. Voor de huidige versie wordt de publieke OpenStreetMap Nominatim-service gebruikt; geen automatische of bulk-aanvragen. Kaart- en adresgegevens worden toegeschreven aan OpenStreetMap-bijdragers.

## Routecontrole

De knop `Controleer route` gebruikt de GPS-coördinaten van vertrek en aankomst en vraagt via de eigen backend een normale autoroute op bij OSRM. De controle is adviserend en blokkeert het opslaan van een rit niet als de routingdienst niet bereikbaar is. De kilometerteller blijft leidend.

## Jaaroverzicht en export

De gebruiker kan per kalenderjaar filteren. De samenvattingskaarten, rittenlijst en CSV-export volgen het gekozen jaar. De jaar-CSV bevat voertuiggegevens, begin- en eindkilometerstand, zakelijke kilometers, privékilometers, totaal en de volledige rittenlijst.

## Correcties en auditlog

Bestaande ritten worden niet hard verwijderd. Een fout wordt gecorrigeerd via `PATCH ./api/rides/<id>` en vereist altijd een reden van minimaal 5 tekens.

Bij iedere correctie bewaart SQLite in `ride_audit`:

- het ritnummer;
- tijdstip van correctie;
- reden;
- de volledige oude versie van de rit;
- de volledige nieuwe versie van de rit.

Voor een correctie wordt de kilometerketen opnieuw gecontroleerd. De nieuwe beginstand moet aansluiten op de vorige rit en de nieuwe eindstand op de volgende rit. Een correctie die de keten verbreekt wordt geweigerd.

De auditlog is leesbaar via `GET ./api/audit` en wordt in de webinterface getoond.

## Dagelijkse back-up

`backup.py` maakt met de SQLite backup-API een consistente kopie van `/var/lib/rittenregistratie/ritten.db`. De tijdelijke kopie wordt met `PRAGMA integrity_check` gecontroleerd voordat deze als geldige back-up wordt gepubliceerd.

Back-ups staan buiten Git in `/var/backups/rittenregistratie/` met bestandsnamen zoals `ritten-20260909T011500Z.db`. Standaard worden back-ups ouder dan 35 dagen verwijderd. Dit is instelbaar met `RITTEN_BACKUP_RETENTION_DAYS`.

Systemd gebruikt:

- `deploy/rittenregistratie-backup.service`
- `deploy/rittenregistratie-backup.timer`

De timer plant dagelijks rond 03:15 lokale systeemtijd met maximaal 10 minuten willekeurige vertraging. `Persistent=true` zorgt dat een gemiste uitvoering na een uitgeschakelde RPi bij de volgende start alsnog wordt ingehaald.

## Huidige status

De kernregistratie is functioneel compleet: centrale opslag, voertuigcontext, GPS/adressen, routecontrole, jaaroverzicht/export en traceerbare correcties. Dagelijkse databaseback-up is in de bron opgenomen. Na installatie op de RPi moet nog expliciet een back-up én een hersteltest worden uitgevoerd. Daarna zijn koppeling vanaf de startpagina en verdere toegangsbeveiliging de belangrijkste vervolgstappen.
