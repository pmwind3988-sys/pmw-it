/**
 * Does this machine's work live somewhere else?
 *
 * A machine that opens its drawings and its ledgers off the server is only as
 * quick as the link to it. Everything the scan already reads — mapped drives,
 * the server folders list, saved server credentials — says that the person
 * works off the network; the network line says how well.
 */

/** A wireless link carrying server work is the bottleneck we see in practice. */
const WIRELESS = /wi-?fi|wireless|wlan/i;

export function serverDependency(device) {
  const mapped = typeof device.mappedDrives === 'number' ? device.mappedDrives : 0;
  const hasFolders = Boolean(device.serverFolders);
  const hasCredentials = Boolean(device.serverCredentials);

  const serverDependent = mapped > 0 || hasFolders || hasCredentials;
  const wireless = WIRELESS.test(device.networkType ?? '');

  if (!serverDependent) {
    return {
      serverDependent: false,
      networkRisk: 'None',
      networkNote: 'Works from its own disk — no server link to depend on',
    };
  }

  if (wireless) {
    const folders = `${mapped || 'its'} server folder${mapped === 1 ? '' : 's'}`;

    // Wireless plus server work is the combination behind most "the file takes
    // forever to open" calls — but only on a machine that could have been
    // plugged in. A laptop is on Wi-Fi because that is what a laptop is for,
    // and calling every one of them critical would bury the desktops sitting
    // three feet from a wall socket.
    if (device.deviceType === 'Laptop') {
      return {
        serverDependent: true,
        networkRisk: 'Wireless',
        networkNote: `Opens ${folders} over Wi-Fi — expected on a laptop, but a dock would be quicker`,
      };
    }

    return {
      serverDependent: true,
      networkRisk: 'Severe',
      networkNote: `Opens ${folders} over Wi-Fi from a desk — put it on a cable`,
    };
  }

  if (!device.networkType) {
    return {
      serverDependent: true,
      networkRisk: 'Unknown',
      networkNote: 'Works off the server, but the scan did not report the network link',
    };
  }

  return {
    serverDependent: true,
    networkRisk: 'Fine',
    networkNote: 'Works off the server over a wired link',
  };
}
