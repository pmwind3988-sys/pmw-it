PMW DEVICE SCAN
===============

Scans one PC for the IT device list. No administrator
rights needed. Nothing is installed on the PC.

EVERY PC
  1. Plug in this drive and open it.
  2. Double-click PMW-Device-Scan.exe
  3. A new PC asks for its location and department.
     A PC scanned before starts on its own.
  4. Wait for "Completed" (about a minute).
     The window closes by itself.

Reports land in the Reports folder on this drive,
named like  [F1 ENGINEERING] AMIR-HP_.txt
Bring them back and drop them on
  IT Portal > Devices > Import scans

IF IT WILL NOT RUN, go down the list
  1. PMW-Device-Scan.exe   the normal way
  2. PMW-Device-Scan.bat   if the program is blocked
                           or will not open. If
                           PowerShell fails too it
                           switches to a basic scan
                           by itself.
  3. PMW-Device-Scan.ps1   if cmd is switched off:
                           right-click > Run with
                           PowerShell
  4. By hand               serial from the sticker,
                           model from Settings > About,
                           add the machine on its page.

SETTINGS (the gear at the bottom right, or hold Shift
while opening the program)
  Save reports to         the Reports folder beside the
                          program, another folder on this
                          drive, or a network folder. If a
                          network folder cannot be reached,
                          the report is saved beside the
                          program instead.
  Start scanning by itself  a PC scanned before starts the
                          moment the program opens.
  Close by itself when done  ten seconds after Completed.
They are kept in PMW-Device-Scan.ini on this drive, and
the .bat and .ps1 follow them too.

WHY IT DOES NOT START WHEN PLUGGED IN
Windows has refused to start programs from a USB
drive by itself since Windows 7. That is a security
rule; leave it on.

WHAT IT READS
Hardware, serial and part numbers, Windows and Office
licences, who Office is signed in as, whether the PC
is joined to PMW Group, and the installed software.
Licence keys are never written, only their last five
characters.

FOR WHOEVER MAINTAINS THIS
  PMW-Device-Scan.bat is the source. The .ps1 is a byte
  copy of it (the batch part is a comment to PowerShell):
  copy it again after every change. The .exe is built
  from exe\ with exe\build.cmd, using the C# compiler
  every Windows has. A label added to one goes into all
  of them and into KNOWN_LABELS in
  src/features/devices/parse/labels.js.
