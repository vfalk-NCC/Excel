# Anteckningar & att göra – 3D-viewer-extension för Trimble Connect

En sidopanel i Trimble Connects 3D-visare där du skriver anteckningar och
att göra-punkter, kopplar dem till objekt och kameravyer i modellen och
exporterar allt till Excel eller Markdown. Använder samma datalager som
4D-planering och kan kopplas mot den.

## Funktioner

- **Två flikar: Anteckningar och Att göra.** En att göra-punkt har också
  förfallodatum, ansvarig och prioritet. En försenad punkt visas i rött,
  och en bock markerar punkten som klar (datumet sparas).
- **Koppla till modellen**: markera objekt i 3D och klicka "🔗 Koppla
  markerade objekt". Kryssa i "Spara aktuell kameravy" om du vill spara
  vyn. 👁 på en post markerar objekten igen och flyttar kameran till den
  sparade vyn, eller zoomar till objekten om ingen vy sparats.
- **Endast markerade objekt**: filtrerar listan på det som är markerat i
  3D just nu.
- **Export** av det som visas (sökning och filter gäller):
  - Excel (`.xlsx`) med flikarna *Att göra* och *Anteckningar*, inklusive
    4D-status/-datum och kopplade objekt.
  - Markdown (`.md`) med kryssrutor för att göra-punkterna.
  - Kopiera som text, till exempel till mejl eller Teams.

## Bubblor i 3D-vyn

Trimbles egna text-markups har fast utseende. Därför ritas varje post som
en egen bild och läggs i modellen med `viewer.addIcon()`. Kryssa i
**💬 Visa bubblor i 3D** ovanför listan.

- **Innehåll:** bubblan visar rubrik, datum och ansvarig, eller författare
  för en anteckning, och 4D-status. Färgen följer typ och status: blå för
  anteckning, orange för att göra, röd för försenad och grön för klar.
  Ramen blir röd om det kopplade 4D-objektet är försenat.
- **Placering:** bubblan sitter ovanför de kopplade objekten, i mitten av
  deras ovansida. Med **📍 Välj punkt för 3D-bubbla** i formuläret kan du
  i stället klicka ut en egen punkt i modellen.
- **Sammanslagning:** flera poster på samma ställe blir en bubbla med
  "+N".
- **Klick:** klickar du på en bubbla i 3D visas och markeras motsvarande
  kort i panelen.
- **Filter:** bubblorna följer aktiv flik, sökning och filter.
- **Zoom:** Trimble ritar ikoner med fast storlek på skärmen. Appen räknar
  därför om storleken när kameran flyttas, så att bubblan krymper när du
  zoomar ut och växer när du zoomar in (upp till 1,5 gånger).
  - **Normalvy:** storleken räknas i förhållande till vyn du har när du
    slår på bubblorna, inte i fasta meter. Det fungerar därför oavsett
    projektets skala och enhet. **↺ Normalstorlek här** gör den nuvarande
    vyn till normalvy.
  - **Nål:** blir bubblan mindre än 60 % av normalstorleken, så att
    texten inte går att läsa, visas en liten färgad nål i stället.
  - **Av:** avståndsanpassningen stängs av under ⚙.
- **Bildformat:** bilderna är kvadratiska (1024 px). Trimble ritar
  ikoner som kvadrater, så en avlång bild trycktes tidigare ihop och gav
  smal, suddig text.
- **Storlek och bild** ställs in under ⚙. Syns inga bubblor, välj *Enkel
  ikon*. Då används färdiga bilder från `callouts/` på GitHub Pages i
  stället för genererade bilder.

## Koppling mot 4D-planering

- **Samma lagring**: med GitHub-token sparas posterna i
  `projects/<projekt-id>/field_notes.json` i det privata repot
  `vfalk-NCC/4D-data`, i samma projektmapp som 4D-planering och
  4D-dashboard. Allt delas då med alla i projektet.
- **Token delas automatiskt**: extensionen hostas på samma origin
  (`vfalk-ncc.github.io`) som 4D-planering. Har du redan angett token där
  används den direkt. Vill du ange en egen görs det under ⚙.
- **Koppla till planerat objekt**: fältet "4D-planering" söker bland
  objekten i `plan_items.json`. Kopplar du markerade objekt som redan är
  planerade föreslås det planerade objektet automatiskt, eftersom båda
  appar matchar på samma externa objekt-ID (IFC GUID). Kortet visar sedan
  4D-objektets status och datum, och det planerade objektets data ändras
  aldrig.
- **💬 Skicka till 4D-planering** lägger posten som en kommentar i
  `plan_item_comments.json`. Den syns då under 💬 på objektet i
  4D-planering.

Utan token körs extensionen i **lokalt läge**. Allt sparas då bara i den
här webbläsaren (`localStorage`), och kopplingen mot 4D-planering är
avstängd.

## Datamodell (`field_notes.json`)

| Fält | Beskrivning |
|---|---|
| `id` | UUID |
| `type` | `note` eller `todo` |
| `title`, `body` | Rubrik och text |
| `done`, `done_at` | Klar-status (att göra) |
| `due_date`, `assignee`, `priority` | Att göra: förfallodatum, ansvarig, `hog`/`normal`/`lag` |
| `plan_item_id` | Id för kopplat objekt i `plan_items.json` (valfritt) |
| `objects` | `[{model_id, object_id, object_name}]`, där `object_id` är det externa ID:t (IFC GUID) |
| `camera` | Sparad kameravy (från `viewer.getCamera()`), annars `null` |
| `author`, `created_at`, `updated_at` | Vem och när |
| `pin` | Egen punkt `{x, y, z}` (meter) för 3D-bubblan, annars `null` |
| `sent_to_4d_at` | När posten senast skickades som kommentar till 4D-planering |

## Publicera och installera

Excel-repot har ingen `main`-branch, så GitHub Pages publiceras direkt från
den branch där `docs/` ligger:

1. Repots **Settings → Pages → Build and deployment**: välj *Deploy from a
   branch*, branch `claude/trimble-connect-app-is9ocn` och mapp `/docs`,
   och klicka **Save**. Efter någon minut svarar
   `https://vfalk-ncc.github.io/Excel/manifest.json`.
2. I Trimble Connect for Browser: öppna projektet och gå till
   **Inställningar → Extensions → Lägg till**. Klistra in
   `https://vfalk-ncc.github.io/Excel/manifest.json` och aktivera
   extensionen.
3. Öppna en modell i 3D-visaren. Extensionen finns i sidopanelen som
   "Anteckningar & att göra".

Varje ny push till branchen publiceras automatiskt igen. Om `docs/` senare
flyttas till en annan branch behöver bara Pages-inställningen i steg 1
pekas om, eftersom manifestets URL är densamma.

**Ny version:** GitHub Pages låter webbläsaren cacha filerna i cirka
10 minuter. Räkna därför upp `APP_VERSION` i `app.js` och `?v=` i
`index.html` vid varje ändring, så att Trimble laddar alla filer i samma
version. Versionen visas i headern bredvid lagringsläget.

**Token:** om GitHub avvisar tokenen (401) provar appen automatiskt
4D-planeringens token. Fungerar ingen visas en varning om att ange en ny
under ⚙.

`github-storage.js` är en identisk kopia av filen i 4D-planering och
4D-dashboard. Uppdatera alla tre samtidigt om den ändras.
