# Rittenregistratie

Deze tool wordt ontwikkeld op de branch `feature/rittenregistratie` en is bedoeld om later te draaien onder:

`https://artsjeroen.ddns.net/tools/rittenregistratie/`

## Ontwikkelstappen

1. Frontend-prototype: rit invoeren, km berekenen, lokale testopslag en CSV-export.
2. Raspberry Pi API + SQLite: centrale en duurzame opslag van ritten.
3. GPS -> adres: locatie op telefoon ophalen en via de eigen backend reverse-geocoden.
4. Routecontrole: optionele routeafstand vergelijken met kilometertellerafstand.
5. Jaaroverzicht + definitieve export.
6. Koppeling vanaf de startpagina en productie-deployment via `main`.

## Belangrijk

- GitHub is de bron van waarheid voor code.
- De ritgegevens zelf horen niet in GitHub.
- Productiedata wordt later op de Raspberry Pi in SQLite opgeslagen, met back-up.
- De frontend moet via HTTPS worden aangeboden voordat telefoongeolocatie betrouwbaar kan worden gebruikt.
- De branch `feature/rittenregistratie` wordt niet naar `main` gemerged voordat de tool is getest.

## Huidige status

Stap 1: frontend-prototype. In deze fase wordt alleen lokale browseropslag gebruikt als testmechanisme. Dit is nadrukkelijk nog niet geschikt als definitieve fiscale administratie.
