import { BarChart, ColumnChart } from '../../../components/ui/Charts';
import { countBy, scansByMonth } from '../stats/deviceStats';
import { ramBucket } from '../deviceFilters';

/**
 * Colour is a signal here, not decoration: anything that means "act on this"
 * is `--it-danger`, "keep an eye on it" is `--it-accent`, "fine" is
 * `--it-good`, and everything neutral is the brand colour.
 */
const token = (name) => `var(${name})`;

const FIT_COLOUR = {
  Critical: token('--it-danger'),
  'Needs Attention': token('--it-accent'),
  Moderate: token('--it-brand'),
  Optimal: token('--it-good'),
  Unknown: token('--it-ink-soft'),
};

const LICENSE_COLOUR = {
  Authentic: token('--it-good'),
  Unlicensed: token('--it-danger'),
  Undefined: token('--it-accent'),
};

const RISK_COLOUR = {
  Critical: token('--it-danger'),
  High: token('--it-danger'),
  Watch: token('--it-accent'),
  OK: token('--it-good'),
  Unknown: token('--it-ink-soft'),
};

const paint = (rows, colourFor) =>
  rows.map((row) => ({ label: row.label, value: row.count, color: colourFor?.(row.label) }));

export default function DeviceCharts({ devices, onFilter }) {
  const select = (key) => (row) => onFilter(key, row.label);

  return (
    <div className="chart-grid">
      <BarChart
        title="Fit for the work"
        blurb="Each machine measured against what its department actually does."
        rows={paint(countBy(devices, (d) => d.fitStatus), (label) => FIT_COLOUR[label])}
        onSelect={select('fit')}
        emptyText="No devices imported yet."
      />

      <BarChart
        title="Office licensing"
        blurb="Authentic means a company product. Undefined means the scan found none."
        rows={paint(countBy(devices, (d) => d.licenseStatus), (label) => LICENSE_COLOUR[label])}
        onSelect={select('license')}
        emptyText="No devices imported yet."
      />

      <BarChart
        title="Workload profile"
        blurb="The bar each machine is held to, taken from its department."
        rows={paint(countBy(devices, (d) => d.personaLabel))}
        onSelect={select('persona')}
        emptyText="No devices imported yet."
      />

      <BarChart
        title="Risk mix"
        blurb="Where the fleet stands. Click a band to see the machines in it."
        rows={paint(countBy(devices, (d) => d.riskLevel), (label) => RISK_COLOUR[label])}
        onSelect={select('risk')}
        emptyText="No devices imported yet."
      />

      <BarChart
        title="Installed RAM"
        blurb="Measured from the memory sticks, not the total Windows reports."
        rows={paint(
          countBy(devices, (d) => ramBucket(d.installedRamGB)),
          (label) => (parseInt(label, 10) <= 8 ? token('--it-accent') : token('--it-brand')),
        )}
        onSelect={select('ram')}
        emptyText="No devices imported yet."
      />

      <BarChart
        title="Laptops and desktops"
        blurb="Read from the motherboard first — the computer name is not reliable."
        rows={paint(countBy(devices, (d) => d.deviceType))}
        onSelect={select('type')}
        emptyText="No devices imported yet."
      />

      <BarChart
        title="Windows"
        blurb="Windows 10 lost security updates on 14 October 2025."
        rows={paint(
          countBy(devices, (d) => d.windowsVersion),
          (label) => (/windows 10|windows [1-9] /i.test(label ?? '')
            ? token('--it-danger')
            : token('--it-brand')),
        )}
        onSelect={select('windows')}
        emptyText="No devices imported yet."
      />

      <BarChart
        title="Storage"
        blurb="A mechanical disk is the cheapest machine-wide speed complaint to fix."
        rows={paint(
          countBy(devices, (d) => d.storageType),
          (label) => (label === 'SSD only' ? token('--it-good') : token('--it-accent')),
        )}
        onSelect={select('storage')}
        emptyText="No devices imported yet."
      />

      <BarChart
        title="Processor age"
        blurb="Obsolete means 6th generation or older, DDR3 memory, or a Pentium."
        rows={paint(
          countBy(devices, (d) => d.cpuAgeBand),
          (label) => ({
            Obsolete: token('--it-danger'),
            Aging: token('--it-accent'),
            Current: token('--it-good'),
          }[label] ?? token('--it-ink-soft')),
        )}
        onSelect={select('cpu')}
        emptyText="No devices imported yet."
      />

      <BarChart
        title="By department"
        blurb="Where the machines are. Click one to see its fleet."
        rows={paint(countBy(devices, (d) => d.department))}
        onSelect={select('department')}
        emptyText="No devices imported yet."
      />

      <ColumnChart
        title="Scans per month"
        blurb="How current the register is."
        columns={scansByMonth(devices).map((row) => ({ label: row.label, value: row.count }))}
      />
    </div>
  );
}
