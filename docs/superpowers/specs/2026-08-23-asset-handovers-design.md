# Asset Handovers — who has what, and when it comes back

**Date:** 2026-08-23
**Status:** Approved design, ready for implementation
**Routes:** `/assets/handover`, `/assets/people/:email`, and additions to `/assets` and `/assets/:id`
**Follows:** `2026-08-23-asset-inventory-design.md`

## 1. Purpose

The register knows what IT owns. It does not know where any of it went.

This section records handing things to people: a new starter's laptop, dock,
mouse and keyboard; a spare monitor lent for a fortnight; three cables out of a
box of twenty. It answers "who has this laptop", "what does Amir hold", "what is
out on loan", and "what is overdue" — and it takes it all back in one action when
somebody leaves.

## 2. Goals

1. Hand several items to one person as a single recorded event.
2. Find people in the company directory rather than typing names.
3. Distinguish **Issued** (it is theirs) from **Borrowed** (it is coming back on
   a date).
4. Handle bulk stock: three of twenty cables to one person, two to another.
5. A page per person listing everything they hold.
6. Take everything back at once, recording the condition each item came back in.
7. Refuse the two things that make a register untrustworthy: a serialised item
   in two places at once, and issuing more than exists.

## 3. Non-goals (v1)

| Excluded | Reason |
|---|---|
| Pre-filling the onboarding form from a person's items | The next project, and much smaller once this exists |
| Emailing or notifying a borrower | A notification system is its own decision; the overdue list is the v1 answer |
| Approval workflow before a handover | IT hands the thing over in person; an approval step would be theatre |
| Assigning to a team or a location rather than a person | A location already exists on the item |
| Transfer straight from one person to another | Return then issue. Two records is the honest description of what happened |
| Signature capture on handover | `SignatureDialog.jsx` exists, but tying it in is its own scope |

## 4. Decisions

### 4.1 Quantity is what you own; what is out is counted separately

A box of twenty cables with three handed out reads **20 owned, 3 out, 17
available** — it does not become a box of seventeen.

Two reasons. "How many do we own" stays answerable, which is the question the
register exists for. And a return becomes `quantityOut - 3` rather than
`quantity + 3`, so a lost return cannot silently inflate the number of cables
the company believes it owns. Arithmetic that only ever moves one derived
figure cannot drift the underlying one.

**Consequence:** `Quantity` on the register keeps its meaning from the previous
project and needs no migration. A new `QuantityOut` column defaults to nothing,
which reads as zero, so every row already saved is correct without being touched.

### 4.2 The handover list is the truth; the item's row carries a readable copy

`IT Asset Handovers` holds one row per item handed over. That is what a return
updates and what the history is read from.

The register row also carries `AssignedTo`, `AssignedToEmail`, `AssignedOn`,
`DueOn` and a `Status` of Assigned or Borrowed — copied, not looked up — so a row
opened directly in SharePoint says who has it without anybody having to join two
lists. Same reasoning as the supplier being copied onto each item in §4.6 of the
register design.

**Where the copy does not apply:** a bulk line can be held by five people at
once, and there is no honest single value for `AssignedTo`. On bulk rows those
fields stay empty and `QuantityOut` carries the answer. The item page says "3 of
20 out, held by 2 people" and lists them from the handover list.

### 4.3 Issued and Borrowed differ by one field

Both are a handover. **Borrowed** carries a due date and therefore appears in the
overdue list; **Issued** does not. Nothing else about them differs, and making
them two mechanisms would double the code for a distinction that is one column.

### 4.4 A basket, so a new starter is one event

Pick the person first, then add items — by searching the register or by scanning
their barcodes with the camera, reusing `useScanner` and the register's own
search. `HandoverId` groups the lines, so four items handed over together can be
reopened as the one thing that happened.

The person comes first deliberately: a basket that does not yet know who it is
for cannot check whether a line is allowed, and a refusal that only arrives at
the end is a refusal nobody can act on while still standing at the desk.

### 4.5 People come from SharePoint's own directory search

`_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`
searches the tenant directory using the SharePoint permissions this app already
holds. Microsoft Graph's `/users` would give job title and department too, but it
needs `User.ReadBasic.All` consented by an admin before anybody can use the
feature at all, and that is a poor trade for two fields.

**Email is the identity**, not the name. "What does Amir have" keys on
`PersonEmail`; the display name is a label that may be spelled two ways.

### 4.6 Two refusals, per line, never per basket

- **A tracked item already out** cannot be issued again. The line is blocked
  with the name of whoever holds it and a link to return it first. A serialised
  thing recorded in two places at once is the failure that makes people stop
  believing the register.
- **More than is available** cannot be issued. Five cables from a box with three
  left is refused with the figure.

Both block the one line. The rest of the basket still goes through, exactly as a
duplicate sticker label does in the register's review grid.

### 4.7 A return records the condition, and a partial return is normal

Taking something back sets its condition as it comes in, so a monitor returned
faulty does not rejoin the shelf as available stock. A bulk line can come back in
pieces — two of the three cables — so a handover row carries `ReturnedQuantity`
and a status of Out, Partly returned or Returned.

Returning a tracked item puts it back to **In stock** and clears the copied
assignment fields on its row. Returning bulk stock decreases `QuantityOut`.

## 5. Architecture

```
src/features/assets/
  people/
    peopleSearch.js        the people-picker request, and normalising its answer
    usePeopleSearch.js     debounced hook over it
  handover/
    basket.js              pure: the basket, its lines, and their defaults
    availability.js        pure: owned / out / available, and who holds a row
    planHandover.js        pure: basket + register -> writes and refusals
    planReturn.js          pure: return lines -> what the two lists become
  sharepoint/
    handoverSchema.js      the list's columns, toListItem / fromListItem
    readHandovers.js       paged read
    writeHandover.js       the writes: handover rows plus the register copies
  useHandovers.js          the one read for handovers
  ui/
    PersonPicker.jsx       directory search box
    BasketLine.jsx         one line of a basket, editable
    HandoverList.jsx       handovers as a table, with a Return button per row
src/pages/
  AssetHandoverPage.jsx    the basket
  AssetPersonPage.jsx      one person and everything they hold
```

`assetSchema.js` gains three columns, `assetViews.js` gains two views,
`provisionAssets.js` gains the handover list. `assetStats.js` gains the out and
overdue figures. Layering is unchanged: `handover/` and `people/` are pure apart
from `peopleSearch.js`, which is the one file that talks to the network.

## 6. Data model

### 6.1 `IT Asset Handovers` (new list)

`Title` is `"<person> — <item>"`, readable and never load-bearing.

| Column | Kind | Notes |
|---|---|---|
| `HandoverId` | text | Groups the lines of one basket |
| `AssetKey` | text | The register row's identity (§4.3 of the register design) |
| `AssetId` | number | The register row's list id, for the return write |
| `ItemTitle` / `Category` | text / choice | Copied, so this list reads on its own |
| `PersonName` | text | Display name at the time of handover |
| `PersonEmail` | text | **The identity.** Everything per-person keys on this |
| `PersonLogin` | text | The claims login name, kept for a later Person column |
| `Quantity` | number | 1 for a tracked item, always |
| `ReturnedQuantity` | number | Supports a partial return (§4.7) |
| `Kind` | choice | Issued, Borrowed |
| `HandoverStatus` | choice | Out, Partly returned, Returned |
| `IssuedOn` / `IssuedOnMYT` | datetime / text | |
| `DueOn` / `DueOnMYT` | datetime / text | Borrowed only |
| `ReturnedOn` / `ReturnedOnMYT` | datetime / text | |
| `ReturnCondition` | choice | The register's condition list |
| `IssuedBy` / `ReturnedBy` | text | |
| `Remarks` | note | |

### 6.2 Added to `IT Asset Register`

| Column | Kind | Notes |
|---|---|---|
| `QuantityOut` | number | How many units are with people (§4.1) |
| `DueOn` | datetime | Tracked borrowed items only |
| `HandoverKind` | choice | Issued, Borrowed — empty when nothing is out |

`AssignedTo`, `AssignedToEmail` and `AssignedOn` were provisioned by the previous
project and start being written here. `Status` starts being set to Assigned or
Borrowed, and back to In stock on return.

### 6.3 Views

On the register: **Out on loan** (`Status` is Assigned or Borrowed). On the
handovers list: **Currently out** and **Overdue** (`HandoverStatus` is not
Returned and `DueOn` is before today) — the overdue view lives here rather than on
the register because a bulk line held by three people has no single due date on
its own row.

## 7. Flows

### 7.1 Handing over

1. `/assets/handover` — search the directory, pick the person.
2. Add lines: search the register, or open the camera and scan. A scanned code
   is matched against `serialNumber`, `assetTag`, `partNumber` and the codes in
   `AdditionalCodes`; no match says so rather than silently adding nothing.
3. Each line takes Issued or Borrowed, a quantity (bulk only), a due date
   (borrowed only) and remarks. Kind and due date can be set once for the whole
   basket and overridden per line.
4. Refusals show against their line as it is added, not at the end (§4.4).
5. **Hand over** writes the handover rows, then updates each register row.

### 7.2 A person

`/assets/people/:email` — everything they hold, what is overdue, and their past
handovers. **Return everything** takes back every open line in one action;
each line can also be returned on its own, with a quantity and a condition.

### 7.3 An item

`/assets/:id` gains a panel: who has it now (or how many are out and with whom),
and every handover it has ever been part of.

## 8. Error handling

| Failure | Behaviour |
|---|---|
| Directory search fails | The box says so and still accepts a typed name and email, so a handover is never blocked by a search outage |
| A line is refused | Blocked with the reason against that line; the rest of the basket saves |
| A handover row writes but the register update fails | Reported per line; the handover list is the truth, so the record is not lost — the item's copied fields are stale until the next handover or a reload, and the result says which rows are affected |
| Two people issue the same laptop at once | The second write finds `Status` already Assigned on re-read and refuses; the register is re-read immediately before the writes for this reason |
| Return of more than is out | Refused with the figure |

## 9. Testing

- `availability` — owned / out / available, tracked vs bulk, missing `QuantityOut`
- `planHandover` — the two refusals, coalescing two lines of one bulk row, the
  tracked quantity pin, per-line kind and due date
- `planReturn` — partial returns, the status transitions, the register updates,
  refusing more than is out
- `basket` — line defaults, batch-level kind and due date inheritance
- `peopleSearch` — normalising the picker's answer, its odd shapes, a failure
- `handoverSchema` — round-trip of every column kind
- `assetStats` — the out and overdue figures

## 10. Acceptance

- [ ] A person can be found in the directory and picked
- [ ] Several items hand over as one event, by search and by scanning
- [ ] A laptop already out cannot be issued again, and says who has it
- [ ] More than is available cannot be issued
- [ ] Three of twenty cables leaves 20 owned, 3 out, 17 available
- [ ] A person's page lists everything they hold and returns it all in one action
- [ ] A return records condition, and a partial return leaves the rest out
- [ ] Overdue borrowed items are listed, in the portal and in SharePoint
- [ ] The item page shows who has it and its full handover history
