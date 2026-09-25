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
| `sent_to_4d_at` | När posten senast skickades som kommentar till 4D-planering |

## Publicera och installera

1. Repots **Settings → Pages**: *Deploy from a branch*, branch `main`,
   mapp `/docs`.
2. I Trimble Connect: projektets **Inställningar → Extensions → Lägg
   till** och ange manifestet
   `https://vfalk-ncc.github.io/Excel/manifest.json`.
3. Aktivera extensionen och öppna den från 3D-visarens sidopanel.

`github-storage.js` är en identisk kopia av filen i 4D-planering och
4D-dashboard. Uppdatera alla tre samtidigt om den ändras.
