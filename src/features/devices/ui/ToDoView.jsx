import { Link } from 'react-router-dom';
import { EmptyState } from '../../../components/ui/Surfaces';
import { describeMove } from '../../assets/machineLink';
import { todoItems } from '../stats/todoItems';
import DepartmentHeatmap from './DepartmentHeatmap';
import FleetMakeup from './FleetMakeup';
import Leaderboards from './Leaderboards';

/**
 * Each item's marker says how bad it is by its shape and colour: a red circle
 * is urgent, an amber triangle a warning, a diamond a suggestion, a grey
 * rounded square upkeep. Colour here only ever means a grade.
 */
const MARK = {
  urgent: { shape: 'circle', glyph: '!' },
  warning: { shape: 'triangle', glyph: '!' },
  suggestion: { shape: 'diamond', glyph: '?' },
  upkeep: { shape: 'square', glyph: '↻' },
};

function Skeleton() {
  return (
    <div className="todo-skeleton" role="status" aria-live="polite">
      Loading the register…
    </div>
  );
}

export default function ToDoView({
  devices, loading, spareCount, offRegister = [], onOpen, onDepartment,
}) {
  const items = todoItems(devices);

  return (
    <div className="todo">
      {spareCount > 0 && (
        <p className="todo-spare">In the IT Stash: {spareCount} ready to hand out</p>
      )}

      <section className="todo-card" aria-labelledby="todo-h">
        <h2 id="todo-h" className="todo-title">
          What needs doing <span className="todo-sub">· worst first</span>
        </h2>
        {loading ? (
          <Skeleton />
        ) : items.length === 0 ? (
          <EmptyState>Nothing needs doing. Every machine is in shape.</EmptyState>
        ) : (
          <ul className="todo-list">
            {items.map((item) => {
              const { shape, glyph } = MARK[item.severity];
              return (
                <li key={item.id} className="todo-row">
                  <span className={`todo-mark todo-mark--${shape} todo-mark--${item.severity}`} aria-hidden="true">
                    {glyph}
                  </span>
                  <span className="todo-text">
                    <span className="todo-item-title">{item.title}</span>
                    <span className="todo-why">{item.why}</span>
                  </span>
                  <span className="todo-count">{item.count}</span>
                  <button
                    type="button"
                    className="todo-show"
                    onClick={() => onOpen?.(item.filter.key, item.filter.value)}
                  >
                    Show machines →
                  </button>
                </li>
              );
            })}
          </ul>
        )}
      </section>

      {offRegister.length > 0 && (
        <section className="todo-card" aria-labelledby="todo-register-h">
          <h2 id="todo-register-h" className="todo-title">
            Not the same as the asset register <span className="todo-sub">· {offRegister.length}</span>
          </h2>
          <ul className="todo-list">
            {offRegister.map(({ device, move }) => (
              <li key={device.id} className="todo-row">
                <span className="todo-mark todo-mark--triangle todo-mark--warning" aria-hidden="true">!</span>
                <span className="todo-text">
                  <span className="todo-item-title">{device.computerName || device.serialNumber || 'Machine'}</span>
                  <span className="todo-why">{describeMove(move, device)}</span>
                </span>
                <Link className="todo-show" to={`/devices/${device.id}`}>Open machine →</Link>
              </li>
            ))}
          </ul>
        </section>
      )}

      <div className="todo-two">
        <DepartmentHeatmap devices={devices} onSelect={onDepartment} />
        <FleetMakeup devices={devices} onOpen={onOpen} />
      </div>

      <Leaderboards devices={devices} />
    </div>
  );
}
