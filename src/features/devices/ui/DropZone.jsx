import { useCallback, useRef, useState } from 'react';
import { Inbox } from '../../../components/ui/Icons';
import Button from '../../../components/ui/Button';

/**
 * Drag and drop OR a file picker. The picker is not optional: dragging an
 * attachment straight out of Outlook does not reliably produce a File, and
 * that is exactly where these reports arrive from.
 */
export default function DropZone({ onFiles, busy }) {
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef(null);

  const stop = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  const handleDrop = useCallback(
    (event) => {
      stop(event);
      setDragging(false);
      const files = [...(event.dataTransfer?.files ?? [])];
      if (files.length) onFiles(files);
    },
    [onFiles],
  );

  return (
    <div
      className={`dz${dragging ? ' dz-active' : ''}${busy ? ' dz-busy' : ''}`}
      onDragEnter={(event) => { stop(event); setDragging(true); }}
      onDragOver={stop}
      onDragLeave={(event) => { stop(event); setDragging(false); }}
      onDrop={handleDrop}
    >
      <Inbox size={28} className="dz-icon" />
      <p className="dz-title">Drop device report files here</p>
      <p className="dz-hint">
        The <code>.txt</code> reports the scan script writes. Drop as many as you like —
        one row per file. Nothing is saved until you have reviewed it.
      </p>
      <Button variant="secondary" onClick={() => inputRef.current?.click()} disabled={busy}>
        Choose files
      </Button>
      <input
        ref={inputRef}
        type="file"
        multiple
        accept=".txt,text/plain"
        className="dz-input"
        onChange={(event) => {
          const files = [...(event.target.files ?? [])];
          if (files.length) onFiles(files);
          event.target.value = '';
        }}
      />
    </div>
  );
}
