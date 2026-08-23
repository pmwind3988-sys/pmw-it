# Native forms — dropping SurveyJS, and rebuilding the asset checklist

**Date:** 2026-08-23
**Status:** Approved design, ready for implementation
**Routes:** `/asset-checklist`, `/it-boarding-form`
**Reference:** the IT ASSET TRACKING FORM the user supplied, used as the field and flow specification

## 1. Purpose

Two forms in this portal are built on SurveyJS. Both are worse for it: the
engine styles itself, so the pages do not look like the rest of the portal;
neither form can be validated without running it; and making it behave has
needed real hacks — a signature button injected as raw HTML and wired up by
`getElementById`, and an interval that spends three seconds hunting the DOM for
buttons to hide.

This replaces the engine with plain React, and rebuilds the asset checklist to
match the reference form exactly.

## 2. Goals

1. `survey-core` and `survey-react-ui` gone from the app entirely.
2. A small shared form kit both forms are built from.
3. `/asset-checklist` follows the reference form's fields, flow and branching.
4. `/it-boarding-form` keeps every field and behaviour it has today.
5. The parts that can be wrong without looking wrong — validation, branching,
   what gets written to SharePoint — become pure functions with tests.

## 3. Non-goals

| Excluded | Reason |
|---|---|
| Connecting either form to the asset register | The register is IT's own; it is not fed by a form a staff member fills in |
| Merging the two forms | They are different things: one is HR raising an event, one is an employee signing for kit |
| Changing what the request form asks | Only its engine changes |
| Reworking `/requests` or the dashboard | Neither list's shape changes in a way they read |
| A general form builder | Two forms, declared in code, is not a product |

## 4. Decisions

### 4.1 The two forms are different things and stay apart

`/it-boarding-form` is HR or a manager raising a **new onboarding or offboarding
event** — a request, feeding `IT Request Form`, which the dashboard and
`/requests` are built on.

`/asset-checklist` is the **employee themselves** signing for what they received
or handed back, feeding `Asset Checklist Form`. It requests nothing.

They overlap in subject and in nothing else. The reference form is the second
one.

### 4.2 A form kit, not a form engine

`src/components/form/` holds the controls: a labelled field, text, number, date,
textarea, select, radio cards, a tick list, repeating rows, and a step wizard.
Each is a controlled React input over a plain object of values.

What it deliberately does not have: a schema interpreter, a rules language, or a
JSON dialect. Two forms declared in code do not need one, and the engine being
replaced is the argument against building another.

### 4.3 The declaration is data; the page is not

Each form's fields, requirements and branching live in a plain module —
`checklistForm.js`, `requestForm.js` — and the validation over them is a pure
function. That is what makes "an OUT checklist must have a signature" and "an
individual request needs at least one item" testable without rendering
anything, which is exactly the part that can be wrong without looking wrong.

### 4.4 The checklist follows the reference form exactly

Two steps, as the reference has them.

**Step 1 — Form Type** (required): `IN` (new joiner provisioning), `OUT` (exit /
offboarding), `INDIVIDUAL REQUEST` (ad-hoc), each with the reference's own
explanatory line.

**Step 2 — the rest:**

| Field | Type | Required |
|---|---|---|
| Employee Name | text | yes |
| Employee No | text | yes |
| Position | text | yes |
| Entity | select — PMW, PCI, PML, WB | yes |
| Date | date, defaulting to today | yes |
| **IN / OUT only:** Asset Checklist | tick list — Laptop, Mouse, Monitor, Keyboard, Speaker, Earphone, Phone & Simcard, Locker Key | no |
| **INDIVIDUAL REQUEST only:** items | repeating item + quantity | at least one item |
| Serial Numbers (Optional) | text | no |
| Other Remarks | text | no |
| Your Signature | signature | yes |

The item list for an individual request is the reference's, which is longer than
the tick list: Laptop, Desktop, Mouse, Monitor, Keyboard, HDMI Cable, VGA Cable,
Speaker, Earphone, Phone & Simcard, Locker Key.

### 4.5 Unlimited request rows, not three

The reference offers exactly three item slots. That is a limit of the tool it
was built in rather than a rule, and somebody needing four things should not
have to submit twice. One row is shown to begin with, and **Add another item**
has no ceiling.

### 4.6 The stored Form Type strings do not change

The form now reads `IN` / `OUT` / `INDIVIDUAL REQUEST`, but what is written stays
`In` / `Out` / `Individual Request` — the values already in the list. Same
meaning, and changing them would orphan every checklist already signed. The
mapping lives in one place, beside the declaration.

### 4.7 Two dates, because they answer different questions

`Date` is what the employee states the handover happened on, and is theirs to
set. `SubmissionDate` stays what it is today: the instant the form was actually
submitted. Keeping both means "when did he get the laptop" and "when was this
signed" cannot be confused, which matters on a form whose whole purpose is being
a signed record.

A new `FormDate` column carries the first. `SubmissionDate` is untouched, so
every existing row keeps meaning what it meant.

### 4.8 Choice columns are reconciled, additively

The request form does not hard-code its dropdowns — it reads them **live from
the SharePoint columns**. So changing Entity to PMW / PCI / PML / WB means
changing the column, not a constant.

`provisionSchema` today only creates missing columns; it never looks at an
existing one's options. It gains that, and **only ever adds**: an option that
records were saved with is never removed, because a value no longer in the
column's list makes those rows unreadable in their own list. The four new
entities are offered from now on; the old three stay valid on old records and
stop being offered on new ones.

### 4.9 The checklist's provisioning moves onto the shared engine

`ensureAssetColumns` in `sharePointService.js` carries the bug already recorded
in `AGENTS.md`: it sends `Choices` with `__metadata: SP.Field`, which the tenant
rejects with *"The property 'Choices' does not exist on type 'SP.Field'"*. It
works today only because its lists predate the bug and would fail outright on a
fresh site.

Since this project rewrites what that function provisions, the checklist moves
onto `features/sharepoint/provision.js` and the broken copy goes. This is the
targeted improvement the work already passes through, not unrelated
refactoring — the request list's own provisioning stays where it is.

## 5. Architecture

```
src/components/form/
  Field.jsx          label, required marker, help text, error
  TextInput.jsx      text / textarea / number / date, one file
  SelectInput.jsx    single choice
  RadioCards.jsx     the big IN / OUT / INDIVIDUAL buttons
  CheckList.jsx      the tick list
  RepeatRows.jsx     add / remove rows
  Wizard.jsx         steps, Previous / Next / Submit, per-step validation
  useFormState.js    values, touched, errors, setField
src/features/forms/
  checklistForm.js   the checklist's fields, per-mode branching, labels
  requestForm.js     the request form's fields
  validate.js        pure: declaration + values -> errors
  toChecklistItem.js pure: values -> the SharePoint payload
  sharepoint/
    checklistSchema.js    columns for `Asset Checklist Form`
    provisionChecklist.js over the shared provisionSchema
src/pages/
  AssetChecklistPage.jsx  rebuilt
  FormPage.jsx            rebuilt
src/styles/forms.css
```

`features/sharepoint/provision.js` gains choice reconciliation (§4.8).
`services/sharePointService.js` loses `ensureAssetColumns` and its list/library
helpers for the checklist; everything the request form uses stays.

## 6. What the checklist writes

`Asset Checklist Form`, as today plus three columns:

| Column | Kind | Notes |
|---|---|---|
| `FormMode` | choice | `In` / `Out` / `Individual Request` (§4.6) |
| `EmployeeName` / `EmployeeNo` / `Position` | text | |
| `Entity` | choice | PMW, PCI, PML, WB — plus whatever is already there (§4.8) |
| `FormDate` | datetime | **new** — the date the employee states |
| `SubmissionDate` | datetime | unchanged — when it was actually submitted |
| `AssetMatrix` | note | the ticked items, for IN and OUT |
| `RequestedItems` | note | **new** — the item + quantity lines, for an individual request |
| `SerialNumbers` | note | **new** |
| `OtherRemarks` | note | **new** |
| `SignatureUrl` | text | unchanged; still the `Signatures` library |

Both item shapes are stored as readable lines rather than JSON — `Laptop x 1`
— because this list is read by people in SharePoint, and `[{"item":"Laptop"…}]`
in a cell is not a record anybody can use.

## 7. Error handling

| Failure | Behaviour |
|---|---|
| A required field is empty | The step refuses to advance and names the fields, all of them at once |
| An individual request with no item | Refused with "add at least one item" |
| Submitted without signing | Refused; the signature is the point of the form |
| SharePoint options fail to load | The request form says so and offers Retry, as today |
| A submit fails | The filled-in form is still there to try again — never cleared on failure |
| The browser is offline | Submit says so rather than failing with a network error |

## 8. Testing

- `validate` — required fields, per-mode requirements, the item-count rule
- `checklistForm` — which fields each mode shows, and the stored-value mapping
- `toChecklistItem` — both item shapes, the two dates, empty values
- `checklistSchema` — the column declaration, internal names, round-trip
- `provision` — choice reconciliation: adds missing, never removes existing

## 9. Acceptance

- [ ] `survey-core` and `survey-react-ui` are gone from `package.json` and the bundle
- [ ] The `.sd-*` and `.survey-light-wrapper` rules are gone from `App.css`
- [ ] `/asset-checklist` asks exactly what the reference form asks, in its order
- [ ] IN and OUT show the tick list; INDIVIDUAL REQUEST shows item + quantity rows
- [ ] An individual request can have more than three items
- [ ] Entity offers PMW / PCI / PML / WB on both forms
- [ ] Signing still works, and still lands in the `Signatures` library
- [ ] `/it-boarding-form` asks exactly what it asked before, still supports
      several employees, and still opens a record with `?edit=`
- [ ] `npm run lint` reports no errors in either page
