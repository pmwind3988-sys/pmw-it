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

export function sortForReview(devices) {
  return [...devices].sort((a, b) => {
    const bHasProblems = issuesFor(b).length > 0 ? 1 : 0;
    const aHasProblems = issuesFor(a).length > 0 ? 1 : 0;
    if (bHasProblems !== aHasProblems) return bHasProblems - aHasProblems;
    return (a.computerName ?? '').localeCompare(b.computerName ?? '');
  });
}
