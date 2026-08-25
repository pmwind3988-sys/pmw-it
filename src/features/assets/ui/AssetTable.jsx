import { Link } from 'react-router-dom';
import { Tag, Boxes } from '../../../components/ui/Icons';
import { formatMYT } from '../../../utils/malaysiaTime';
import { isTracked } from '../assetKinds';
import { available, out } from '../handover/availability';
import { unitsOf, perItem, isBlankUnit } from '../units';

/**
 * The register, as a list.
 *
 * A quantity is shown as "× 20" rather than as a bare number in a column
 * nobody reads the header of, because the difference between one mouse and
 * twenty is the whole point of a bulk row.
 */

/** A tracked row is one thing, so "available" is a name rather than a number. */
function withWhom(asset) {
  return asset.assignedTo ? `With ${asset.assignedTo}` : 'Out';
}

/**
 * A bulk line's serials and labels are on its items, and a column is not the
 * place to list twenty of them. It shows the first and says how many more
 * there are — enough to recognise the row, and the item detail has the rest.
 */
function Codes({ asset }) {
  if (isTracked(asset.trackingMode)) {
    return (
      <>
        {asset.serialNumber || '—'}
        {asset.assetTag && <span className="as-tag"><Tag size={11} />{asset.assetTag}</span>}
      </>
    );
  }

  const units = unitsOf(asset).filter((unit) => !isBlankUnit(unit));
  if (!units.length) return <span className="as-sub">per item — none recorded</span>;

  const [first] = units;
  const more = units.length - 1;

  return (
    <>
      {first.serialNumber || '—'}
      {first.assetTag && <span className="as-tag"><Tag size={11} />{first.assetTag}</span>}
      {more > 0 && <span className="as-sub">+{more} more recorded</span>}
    </>
  );
}

/**
 * Conditions across everything on the row. One value reads as itself; a line
 * where the items disagree says so rather than picking a winner.
 */
function conditionOf(asset) {
  const stated = perItem(asset, 'condition').filter((entry) => entry.value);
  if (!stated.length) return '—';
  if (stated.length === 1) return stated[0].value;

  return stated
    .sort((a, b) => b.count - a.count)
    .map((entry) => `${entry.count} ${entry.value.toLowerCase()}`)
    .join(', ');
}

export default function AssetTable({ assets }) {
  return (
    <div className="as-table-wrap">
      <table className="as-table">
        <thead>
          <tr>
            <th>Item</th>
            <th>Category</th>
            <th>Serial / label</th>
            <th>Qty</th>
            <th>Available</th>
            <th>Condition</th>
            <th>Where</th>
            <th>Arrived</th>
          </tr>
        </thead>
        <tbody>
          {assets.map((asset) => (
            <tr key={asset.id}>
              <td>
                <Link to={`/assets/${asset.id}`} className="as-link">
                  {asset.title || asset.model || 'Untitled item'}
                </Link>
                {asset.supplier && <span className="as-sub">from {asset.supplier}</span>}
              </td>
              <td>
                <span className="as-chip">
                  {isTracked(asset.trackingMode) ? null : <Boxes size={12} />}
                  {asset.category || '—'}
                </span>
              </td>
              <td className="as-mono"><Codes asset={asset} /></td>
              <td className="as-qty">{isTracked(asset.trackingMode) ? '1' : `× ${asset.quantity ?? 1}`}</td>
              <td className="as-qty">
                {/* What is owned never moves when something is handed out, so
                    this is the figure that answers "can I give one away". */}
                {isTracked(asset.trackingMode)
                  ? (out(asset) ? withWhom(asset) : 'Yes')
                  : available(asset)}
              </td>
              <td>{conditionOf(asset)}</td>
              <td>{asset.location || '—'}</td>
              <td className="as-when">
                {/* The stored readable copy is preferred: it was written in
                    Malaysia time at the moment it happened, which is the answer
                    somebody wants — not this browser's idea of the timezone. */}
                {asset.arrivedOnMYT
                  || (asset.arrivedOn ? formatMYT(asset.arrivedOn, 'datetime12') : '—')}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
