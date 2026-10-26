using System.Reflection;
using System.Runtime.InteropServices;

// What Windows shows under the file's Properties > Details, and what antivirus
// reputation checks read. An unsigned program with all of these blank looks
// like something thrown together to hide; filled in, it says plainly what it is.
[assembly: AssemblyTitle("PMW Device Scan")]
[assembly: AssemblyDescription("Reads this PC's hardware, licences and installed software into a text report for the PMW IT device list. Changes nothing on the PC and sends nothing anywhere.")]
[assembly: AssemblyCompany("PMW Group")]
[assembly: AssemblyProduct("PMW Device Scan")]
[assembly: AssemblyCopyright("PMW Group IT")]
[assembly: AssemblyVersion("2.0.0.0")]
[assembly: AssemblyFileVersion("2.0.0.0")]
[assembly: ComVisible(false)]
