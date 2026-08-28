import { useState } from 'react';
import { typedQuantity, settledQuantity } from '../quantityText';

/**
 * The "how many?" box, which can be emptied.
 *
 * It used to snap back to 1 the instant it was cleared, so changing a 1 to a 3
 * meant typing over a number that kept reappearing under the cursor — and
 * coming out with 13. The box now holds whatever is being typed, including
 * nothing, and only tells the row about a number when there is one. Leaving it
 * empty puts the previous count back rather than inventing a 1.
 *
 * What is being typed is remembered from the first keystroke rather than from
 * the moment the box is focused. Focus is not a reliable signal — a box filled
 * in by a phone's autofill, by a barcode wedge, or by anything that writes
 * without focusing first would never start holding its own text, and would go
 * back to fighting the person typing in it.
 */
export default function QuantityInput({ value, onCommit, ...rest }) {
  // `null` means "nothing is being typed, show the count the row has".
  const [text, setText] = useState(null);
  const shown = text === null ? String(value ?? '') : text;

  return (
    <input
      type="number"
      min="1"
      inputMode="numeric"
      value={shown}
      // Tapping it selects what is there: the usual reason for touching this
      // box on a phone is to replace the number, not to add a digit to it.
      onFocus={(event) => event.target.select()}
      onChange={(event) => {
        setText(event.target.value);
        const number = typedQuantity(event.target.value);
        if (number !== null) onCommit(number);
      }}
      onBlur={() => {
        // An empty box left empty is not a change of mind about the count.
        onCommit(settledQuantity(text, value));
        setText(null);
      }}
      {...rest}
    />
  );
}
