import { AlertTriangle, RefreshCw } from './Icons';

export function Card({ children, className = '', ...rest }) {
  return (
    <div className={`ui-card ${className}`.trim()} {...rest}>
      {children}
    </div>
  );
}

/**
 * The inline error state, used identically on every screen: the message wraps
 * rather than squashing, and Retry only appears when there is something to
 * retry.
 */
export function ErrorBanner({ message, onRetry }) {
  return (
    <div className="ui-error" role="alert">
      <span style={{ display: 'flex', gap: 8, alignItems: 'flex-start', minWidth: 0 }}>
        <AlertTriangle size={15} style={{ marginTop: 1, flexShrink: 0 }} />
        <span style={{ minWidth: 0 }}>{message}</span>
      </span>
      {onRetry && (
        <button type="button" onClick={onRetry}>
          <RefreshCw size={13} /> Retry
        </button>
      )}
    </div>
  );
}

export function EmptyState({ children }) {
  return <div className="ui-empty">{children}</div>;
}
