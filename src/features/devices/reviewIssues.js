/**
 * The problems worth pulling to the top of the review table. Each one is
 * phrased as something the reader can act on, not as a code.
 */
export function issuesFor(device) {
  const issues = [];

  if (device.scanComplete === false) issues.push('Scan incomplete — most fields are empty');
  if (device.deviceType === 'Unknown') issues.push('Device type could not be determined');

  if (device.ramDiscrepancy) {
    issues.push(
      `Reports ${device.reportedRamGB} GB usable of ${device.installedRamGB} GB installed `
        + '— the GPU reserves the rest',
    );
  }

  for (const unknown of device.unknownLabels ?? []) {
    issues.push(`New field found in the report: ${unknown.label}`);
  }

  if (!device.owner) issues.push('No owner could be resolved');

  return issues;
}

/**
 * Rows waiting on a replacement answer first -- Save cannot go ahead without
 * them -- then rows with problems, then by name.
 */
export function sortForReview(devices, first = new Set()) {
  return [...devices].sort((a, b) => {
    const aFirst = first.has(a.sourceFileName) ? 1 : 0;
    const bFirst = first.has(b.sourceFileName) ? 1 : 0;
    if (aFirst !== bFirst) return bFirst - aFirst;
    const bHasProblems = issuesFor(b).length > 0 ? 1 : 0;
    const aHasProblems = issuesFor(a).length > 0 ? 1 : 0;
    if (bHasProblems !== aHasProblems) return bHasProblems - aHasProblems;
    return (a.computerName ?? '').localeCompare(b.computerName ?? '');
  });
}
