# Release notes

Datum: 5 september 2026  
Branch: `release`  
Bron: `main`

Deze branch is bedoeld als release-overzicht. De functionele websitecode is afkomstig van `main`; in deze releasebranch is alleen dit release-notesbestand toegevoegd.

## Startpagina

### Wijzigingen

- Zijbalk opgeschoond en aangepast; WhatsApp vervangt Gemini.
- Linkstructuur voor wetgeving bijgewerkt, waaronder de Markttoezichtverordening en het Warenwetbesluit bestuurlijke boeten.
- Sectie **Arbo Wetgeving** hernoemd naar **Overige Wetgeving** en een link naar de Algemene wet bestuursrecht toegevoegd.
- Mobiele iconenbalk compacter gemaakt.
- Opmaak en typografie van normen en harmonisatielinks rustiger en consistenter gemaakt.
- Ongebruikte websitebestanden verwijderd.
- Startpagina hersteld naar de gewenste versie na een tijdelijke testwijziging.
- Lotus-snelkoppeling naar de ademhalingstool toegevoegd.
- GitHub-bewerkknop als los zwevend element geplaatst.
- Navigatie, thema en utility rail voorbereid op gedeeld gebruik met de ademhalingstool.
- Overgang van `/main/index.html` naar `/ademhalingstool/index.html` vloeiender gemaakt.

### Relevante paden

- `main/`
- gedeelde styles/scripts die door startpagina en ademhalingstool worden gebruikt

## Ademhalingstool

### Nieuwe functionaliteit

- Nieuwe pagina voor begeleide ademhaling toegevoegd.
- Vier fasen instelbaar: inademen, vasthouden, uitademen en vasthouden.
- Vooraf ingestelde ademhalingsritmes toegevoegd, waaronder 4-7-8 en box breathing.
- Maximale fasewaarde uitgebreid zodat langere patronen mogelijk zijn.
- Timinglogica en afronding van sessies toegevoegd.
- Fasegeluiden toegevoegd en later zachter/vriendelijker gemaakt.
- Instellingen worden lokaal onthouden.
- Standaard aantal herhalingen staat op 20.
- Tijdens een actieve sessie wordt geprobeerd het scherm wakker te houden.

### Interface en gebruik

- Visuele stijl afgestemd op de startpagina, inclusief licht/donker thema.
- Dezelfde utility rail, links, iconen en afmetingen als op de startpagina toegepast.
- Homeknop toegevoegd en navigatie terug naar hetzelfde tabblad hersteld.
- Mobiele portretweergave geoptimaliseerd zodat de tool beter binnen één scherm past.
- Mobiele geluidsbediening vervangen door een compact icoon rechtsboven.
- Mobiele audio-unlock toegevoegd voor browsers die geluid pas na een gebruikersactie toestaan.
- Desktopindeling aangepast naar bediening links en een grotere animatie rechts.
- Ademhalingsbal op desktop groter gemaakt en verticaal beter gecentreerd.
- Tekst boven de ademhalingsbal verwijderd om de animatie rustiger te maken.
- Samenvatting vereenvoudigd: `Per cyclus` verwijderd; totale tijd en resterende tijd samengevoegd.
- Handmatige fasevelden worden alleen getoond bij **Eigen ritme**.
- Presetlogica behouden zonder overbodige toelichting in de interface.

### Relevante paden

- `ademhalingstool/`
- gedeelde styles/scripts voor thema, achtergrond en utility rail

## Testpunten voor deze release

### Startpagina

- Open `/main/index.html` op desktop en mobiel.
- Controleer licht/donker thema.
- Controleer alle zijbalkiconen en externe links.
- Controleer **Overige Wetgeving** en de link naar de Algemene wet bestuursrecht.
- Open de ademhalingstool via de lotusknop en controleer de vloeiende overgang.

### Ademhalingstool

- Test een standaard preset en **Eigen ritme**.
- Controleer dat de handmatige velden alleen bij **Eigen ritme** zichtbaar zijn.
- Start een sessie met 20 herhalingen en controleer de teruglopende tijd.
- Test pauzeren/stoppen/afronden indien beschikbaar in de huidige interface.
- Test geluid op desktop en mobiel; controleer het geluidsicoon rechtsboven.
- Test portrait mobiel zonder onnodig scrollen.
- Test desktop: grote, verticaal gecentreerde ademhalingsbal.
- Controleer dat het scherm tijdens een actieve sessie niet ongewenst in slaap valt wanneer de browser dit ondersteunt.

## GitHub en deployment

- `main` blijft de bron van waarheid voor ontwikkeling.
- `release` is aangemaakt vanaf de actuele `main`-stand van 5 september 2026.
- In `release` is alleen `RELEASE_NOTES.md` toegevoegd; er zijn geen websitebestanden aangepast.
- Er is met deze wijziging **geen RPi-deployment uitgevoerd**.
- Voor deployment naar de Raspberry Pi: deploy expliciet vanaf de gewenste, geteste GitHub-ref en verifieer daarna zowel `/main/index.html` als `/ademhalingstool/index.html` op het apparaat.
