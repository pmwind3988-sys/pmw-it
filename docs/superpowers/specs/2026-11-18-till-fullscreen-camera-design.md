# Till: full-screen camera, tap to add

**Why:** a box carries four or five barcodes and the till took whichever it
read. The person holding the box should decide which number goes on the list.

**What it does (agreed 2026-11-18):**
- "Use the camera" fills the whole screen. ✕ (or Escape) returns to the till;
  a corner badge says how many things are on the list.
- Every number the camera reads appears in a card drawn over the picture.
  Nothing is added on its own: a number goes on the list only when tapped.
  Several can be tapped from one box.
- A tapped number shows "✓ added" while still in view and cannot be tapped
  again. A number the list already holds (a serial on the receipt, a serial
  already in the run) is never shown.
- A number not tapped leaves the card a few seconds after the camera stops
  seeing it.
- Applies to every till tab that has a camera and to the "scan more of the
  same" sheet. Questions a tapped number raises (New item) open over the
  camera as now.

**Approaches considered:**
1. *Keep the read gate, offer every accepted code as a choice* — still guesses
   which code was "meant" and asks one at a time; rejected.
2. *A pick list owned by the camera* (chosen) — the camera collects what it
   reads, the page only says what to skip, how to describe a code, and what a
   tap does. One component serves the till and the serial run.
3. *Labels drawn on each barcode's position in the picture* — the user asked
   for a card over the view, not markers; positions jitter on phones.

**Builder's calls:** a code must be read twice within 1.5 s to appear (keeps
misreads out); it is forgotten 4 s after it was last seen, after which it can
be shown and tapped again (the second identical box of mice counts up);
"never shown" is decided by a dry run of the mode's own scan rule returning
"already on the receipt". The card holds at most six codes, oldest first so
rows do not jump under a finger.
