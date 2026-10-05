/** The one "this is working" mark. Inherits the text colour of whatever holds it. */
export default function Spinner({ size = 14 }) {
  return <span className="ui-spinner" style={{ width: size, height: size }} aria-hidden="true" />;
}
