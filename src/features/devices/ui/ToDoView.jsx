import { EmptyState } from '../../../components/ui/Surfaces';
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

export default function ToDoView({ devices, loading, spareCount, onOpen, onDepartment }) {
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

      <div className="todo-two">
        <DepartmentHeatmap devices={devices} onSelect={onDepartment} />
        <FleetMakeup devices={devices} onOpen={onOpen} />
      </div>

      <Leaderboards devices={devices} />
    </div>
  );
}
