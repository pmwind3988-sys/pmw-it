# Asset Inventory — camera-scanned register of what IT owns

**Date:** 2026-08-23
**Status:** Approved design, ready for implementation planning
**Routes:** `/assets`, `/assets/scan`, `/assets/batch/:id`, `/assets/:id`

## 1. Purpose

IT buys things — laptops, monitors, printers, docks, bags of mice and cables — and
today there is no record of what arrived, when, from whom, or where it went. The
purchase order is in somebody's mail and the serial number is on a box in the store
room.

This section is the register of what the company owns, and the fast way to fill it.
A delivery is recorded by pointing a phone camera at the boxes: barcodes become draft
rows, the purchase details are entered once for the whole delivery, and the batch is
reviewed and saved to SharePoint in one deliberate action. Nothing about that requires
a signal in the store room.

## 2. Goals

1. Add, edit and remove items in a register backed by SharePoint.
2. Scan barcodes with a phone camera — one item with several codes on it, or many
   items in a sweep — without ever counting the same code twice.
3. Keep working with no connection: a scanning session lives on the phone until it
   is reviewed and saved.
4. Record purchase details once per delivery and let them flow to every item in it.
5. Photograph the PO once per delivery and each item individually.
6. Handle both kinds of thing IT owns: units with serial numbers, and bulk stock
   counted by quantity.
7. Track the printed sticker label on a machine as a unique code that cannot be
   given to two items.

## 3. Non-goals (v1)

| Excluded | Reason |
|---|---|
| Looking specs up on the internet from a product name | Needs a server-side search this app does not have; its own project |
| Assigning an item to an M365 person | Its own project; the columns are provisioned now so it is not a migration |
| Pre-filling the onboarding / offboarding form from marked items | Depends on assignment; its own project |
| Linking an asset to its `/devices` row | Deliberate (§4.8) |
| Purchase cost, depreciation, warranty expiry | Not asked for; adding a field later is cheap |
| A USB barcode gun | It types like a keyboard, so the manual-entry field already accepts it |
| Editing the register from SharePoint and syncing back | SharePoint is the store, the portal is the editor |

## 4. Decisions

### 4.1 A scanning session is an object, not a connection

Scanning produces a **batch** held in the browser's own storage: the codes, the photos
as blobs, and the purchase details. Nothing reaches SharePoint until the batch is
reviewed and saved.

This is the same shape as `/devices` (drop files → review grid → save) and it is what
makes offline honest rather than bolted on. A queue that retries failed writes would
have to reason about half-written deliveries; a batch that has simply not been saved
yet has no such state.

**Consequence:** an unreviewed batch is invisible to everyone but the phone that holds
it. `/assets` therefore carries a banner counting unsaved batches that cannot be
dismissed, only opened.

### 4.2 Two kinds of item, decided by category

A **tracked** item is one physical unit: it has a serial number, its own photo, its
own sticker label, its own history. A **bulk** line is a model with a quantity that
goes up and down.

Which one an item is follows from its category — laptops, desktops, monitors, printers,
docks and phones are tracked; mice, keyboards, cables, adapters and consumables are
bulk — so it is never a question answered twice or answered differently by two people.
`assetKinds.js` holds that mapping, and the review grid lets a row be switched by hand
for the case the mapping gets wrong (a serialised keyboard, an untracked spare monitor).

### 4.3 Upsert by identity, so a re-scan is never a duplicate

Every asset carries a normalised `AssetKey`:

- tracked → `serial:<manufacturer>|<serial>`
- bulk → `bulk:<category>|<manufacturer>|<model>`

Scanning a laptop that is already in the register **updates** that row. Scanning
another bag of the same mice **adds to** the existing quantity rather than creating a
second line. This is the dedupe that matters — the in-session one (§4.4) only stops
the same box being counted twice in one sweep.

A tracked item with no serial falls back to `tag:<assetTag>`, and failing that gets a
generated `local:<uuid>` — recorded so that "this row has no stable identity" is
visible rather than silently producing duplicates forever.

### 4.4 Two scan modes, chosen before the camera opens

**One item** pools every code seen into a single draft and works out which code is
which (§4.5). You confirm, and the pool resets for the next box.

**Many items** makes each newly seen code its own draft.

The mode is picked up front because the app cannot tell the difference between "two
codes on one box" and "two boxes" from the video alone, and guessing wrong silently
corrupts the count.

Within a session every accepted code is remembered for the whole session, so a barcode
that stays in frame cannot be counted again. Re-reading a remembered code is
acknowledged out loud — a distinct sound and a grey "already scanned" flash — because
silence is indistinguishable from the scanner not working.

### 4.5 A code's meaning is guessed from its shape, and always editable

Manufacturers do not label which barcode is the serial. `classifyCode.js` scores each
code against patterns — an explicit `S/N`, `SN:` or `PN:` prefix wins outright; a MAC
address is unmistakable; the longest mixed alphanumeric is the serial by default and
a shorter code with a leading letter-run is the part number.

Every guess lands in a field the review grid shows as *guessed* and lets you retype,
exactly as the device import flags derived values. Codes that were not claimed by any
field are kept verbatim in `AdditionalCodes` rather than thrown away.

### 4.6 Purchase details belong to the delivery, not the item

Supplier, PO number, arrival date and time, and the PO photo attach to the batch and
are copied onto every asset saved from it. A delivery arrives as a delivery; typing
the supplier thirty times is how a register stops being filled in.

They are copied onto the row rather than looked up through a reference, so one row read
in SharePoint is complete on its own — the same reasoning as raw-answer-beside-derived
in the device list. The batch keeps its own row in `IT Asset Batches` so a delivery can
still be opened as a whole.

Any of them can be overridden on an individual row before saving, for the case where
one line on the delivery note came from a different supplier.

### 4.7 Label codes are unique, and the register knows what is unlabelled

`AssetTag` is checked for uniqueness at review time against the register and again at
save. A clash blocks that one row with a link to the item already wearing the code; the
rest of the batch still saves. A saved SharePoint view lists tracked items with no tag,
so "what still needs a sticker" is one click.

### 4.8 Separate from `/devices` for now

`/devices` answers "what is inside this machine and is it healthy", built from scan
reports. This answers "what do we own, where is it, who bought it". They will eventually
meet on the serial number, but merging a working section into a new one doubles this
project. `src/features/assets/` is a sibling of `devices/` and imports nothing from it
except `../datastudio/time/malaysiaTime.js`.

### 4.9 Native barcode decoding where it exists, a ponyfill where it does not

Chrome on Android decodes barcodes natively through `BarcodeDetector`, several per
frame, fast. Safari has no such thing. `scan/detector.js` returns the native detector
when present and otherwise dynamically imports a WebAssembly ponyfill — dynamically, so
an Android phone never downloads it.

Both are addressed through one interface that takes a video frame and returns
`{ rawValue, format }`, which is also what the tests fake. No camera code is under test;
the decision of what to do with a decoded code is, and that is the part that can be
wrong without looking wrong.

## 5. Architecture

```
src/features/assets/
  assetKinds.js          category → tracked/bulk, and the category list
  identity.js            AssetKey derivation, normalisation
  scan/
    detector.js          BarcodeDetector or ponyfill (the only impure file here)
    classifyCode.js      which code is serial / part / MAC / unknown
    scanSession.js       pure: codes + mode + seen-set → drafts
  draft/
    draftAsset.js        a draft row, its defaults and its validation
    batch.js             pure: batch shape, purchase-detail inheritance, merges
  store/
    assetDb.js           IndexedDB: batches and photo blobs
  sharepoint/
    assetSchema.js       columns, toListItem / fromListItem
    assetViews.js        the saved views
    provisionAssets.js   lists, columns, views, photo library
    readAssets.js        paged read of the register
    planSave.js          pure: batch + register → inserts / updates / conflicts
    saveBatch.js         provisioning, photo upload, writes, change log
    uploadPhoto.js       binary upload to the document library
    updateAsset.js       edit and delete one row
  stats/assetStats.js    the figures on /assets
  ui/                    ScanScreen, ReviewGrid, AssetTable, BatchBanner, PhotoInput
  useAssets.js           the one SharePoint read for the section
src/pages/
  AssetsPage.jsx  AssetScanPage.jsx  AssetBatchPage.jsx  AssetDetailPage.jsx
src/styles/assets.css
```

Layering follows the house rule: `scan/` and `draft/` know nothing about SharePoint,
`sharepoint/` imports no React, `ui/` is the only part that does.

## 6. Data model

### 6.1 `IT Asset Register`

`Title` is a readable name — `"Dell P2422H — CN0ABC123"` for tracked, `"Logitech B100"`
for bulk — built at save. Identity lives in `AssetKey`, never in `Title`.

| Column | Kind | Notes |
|---|---|---|
| `AssetKey` | text | §4.3; the upsert key |
| `Category` | choice | Laptop, Desktop, Monitor, Printer, Docking Station, Phone, Keyboard, Mouse, Cable, Adapter, PC Part, Network, Accessory, Other |
| `TrackingMode` | choice | Tracked, Bulk |
| `Manufacturer` / `Model` | text | |
| `SerialNumber` / `PartNumber` | text | |
| `AdditionalCodes` | note | every scanned code not claimed by a field |
| `AssetTag` | text | the printed sticker; unique when present (§4.7) |
| `Quantity` | number | always 1 for tracked |
| `Condition` | choice | New, Good, Fair, Faulty, Retired |
| `Status` | choice | In stock, Assigned, Borrowed, In repair, Retired, Disposed |
| `Location` | text | store room, shelf, site |
| `Remarks` | note | |
| `SpecSummary` | note | typed now, auto-filled by a later project |
| `Supplier` / `PoNumber` | text | inherited from the batch (§4.6) |
| `ArrivedOn` | datetime | `DisplayFormat: 1` — the time of day matters |
| `ArrivedOnMYT` | text | readable AM/PM, as the device list does |
| `PurchasedOn` | datetime | optional; the invoice date where it differs |
| `BatchId` / `BatchTitle` | text | which delivery this came from |
| `PhotoUrl` / `PoPhotoUrl` | text | into `IT Asset Photos` |
| `ScanSource` | choice | Camera, Manual |
| `GuessedFields` | note | which values were guessed rather than read (§4.5) |
| `ManualFields` | note | which were typed by hand, and so outrank a re-scan |
| `AssignedTo` / `AssignedToEmail` | text | provisioned, no UI in v1 |
| `AssignedOn` | datetime | provisioned, no UI in v1 |
| `AddedOn` / `AddedOnMYT` / `AddedBy` | datetime / text / text | |

### 6.2 `IT Asset Batches`

One row per delivery: `Title` = batch reference, plus `Supplier`, `PoNumber`,
`ArrivedOn`, `ArrivedOnMYT`, `PoPhotoUrl`, `ItemCount`, `SavedOn`, `SavedBy`, `Remarks`.

### 6.3 `IT Asset Changes`

The device change log's columns exactly (`FieldName`, `OldValue`, `NewValue`,
`ChangedOn`, `ChangedOnMYT`, `ChangedBy`, `ChangeType`), keyed by `Title` = asset key.

Tracked fields: category, manufacturer, model, serialNumber, partNumber, assetTag,
quantity, condition, status, location, supplier, poNumber. `AdditionalCodes` and the
photo URLs are excluded — they churn without meaning anything.

### 6.4 `IT Asset Photos`

A document library. Files are named `<assetKey-slug>-<timestamp>.jpg`; PO scans go to a
`po/` folder under it. Photos are captured as JPEG at a longest edge of 1600px, which
keeps a full delivery's photos inside a phone's storage budget and still reads a label.

## 7. Flows

### 7.1 Scan → review → save

1. `/assets` → **Scan a delivery**. The purchase sheet is first: supplier, PO number,
   arrival date and time (defaulted to now, Malaysia time), and an optional PO photo.
   All of it is skippable and editable later.
2. Choose *one item* or *many items*, then the camera opens.
3. Each accepted code appends a draft to the strip along the bottom. Tapping a draft
   opens it to set category, quantity and photo without leaving the camera.
4. **Done** ends the session; the batch persists.
5. `/assets/batch/:id` — the review grid: every row editable, guessed values flagged,
   duplicates against the register shown as "will update", tag clashes shown as
   blocking.
6. **Save** provisions if needed, uploads photos, writes rows, logs changes, and reports
   per-row success exactly as the device import does.

### 7.2 Manual entry

The same review grid with an **Add row** button, reachable from `/assets` without ever
opening a camera. This is also the path a USB barcode gun uses — it types into the
serial field.

### 7.3 Editing and removing

`/assets/:id` shows one asset in full and edits it in place through `updateAsset.js`,
recording touched fields in `ManualFields` so a later re-scan does not undo the
correction. Removal is a confirmed delete of the SharePoint row, with a change-log entry.

## 8. Error handling

| Failure | Behaviour |
|---|---|
| Camera permission refused | The screen explains it and offers manual entry; the batch is unaffected |
| No barcode decoder at all | Same — manual entry, stated plainly |
| Browser storage full | The existing `StorageFullError` pattern: say which batches are taking room |
| Offline at save | Save is blocked with "you are offline"; the batch is untouched and retriable |
| One row fails to write | The rest still save; the batch keeps only the failed rows so Save can be pressed again |
| Photo upload fails | The row still saves without its photo, flagged in the result |
| Tag clash | That row alone is blocked, with a link to the offender |

## 9. Testing

Vitest, as the rest of the repo: pure modules carry the weight.

- `classifyCode` — prefixed codes, MACs, a real Dell/HP/Lenovo label's code set, ties
- `scanSession` — dedupe across frames, both modes, re-scan acknowledgement
- `identity` — key derivation, the no-serial fallbacks, normalisation
- `batch` — purchase inheritance and per-row override
- `planSave` — insert vs update vs quantity-add vs tag clash
- `assetSchema` — round-trip of every column kind, `false` and `0` surviving
- `assetDb` — via `fake-indexeddb`, as the Data Studio store is tested
- `assetStats` — the figures on the page

## 10. Acceptance

- [ ] `/assets` lists the register, searchable and filterable, with unsaved batches banner
- [ ] A delivery can be scanned, reviewed offline and saved when back online
- [ ] Several barcodes on one box become one item with the serial identified
- [ ] The same code cannot be counted twice in a session, and says so when re-read
- [ ] A re-scanned laptop updates its row; a re-scanned bulk model adds to its quantity
- [ ] A duplicate sticker label is refused with a link to the item that has it
- [ ] Items can be added, edited and removed by hand, with photos
- [ ] Lists, columns, views and the photo library provision themselves on first save
