/**
 * The one button. Variants and sizes are class names on `.ui-btn` — see
 * `src/styles/shell.css`.
 */
export default function Button({
  children,
  variant = 'primary',
  size = 'md',
  icon: Icon,
  className = '',
  ...props
}) {
  return (
    <button className={`ui-btn ui-btn-${size} ui-btn-${variant} ${className}`.trim()} {...props}>
      {Icon && <Icon size={14} />}
      {children}
    </button>
  );
}
