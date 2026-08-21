import { useCallback, useRef, useState } from 'react';
import { Inbox } from '../../../components/ui/Icons';
import Button from '../../../components/ui/Button';

/**
 * Drag and drop OR a file picker. The picker is not optional: dragging an
 * attachment straight out of Outlook does not reliably produce a File, and
 * that is exactly where these reports arrive from.
 */
export default function DropZone({ onFiles, busy, compact = false }) {
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef(null);
  // Entering a child fires dragleave on the parent, so a plain boolean would
  // clear the highlight while the pointer is still over the zone. Counting
  // enter/leave pairs keeps it lit until the pointer actually leaves.
  const depth = useRef(0);

  const stop = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  const handleDrop = useCallback(
    (event) => {
      stop(event);
      depth.current = 0;
      setDragging(false);
      const files = [...(event.dataTransfer?.files ?? [])];
      if (files.length) onFiles(files);
    },
    [onFiles],
  );

  return (
    <div
      className={`dz${compact ? ' dz-compact' : ''}${dragging ? ' dz-active' : ''}${busy ? ' dz-busy' : ''}`}
      onDragEnter={(event) => { stop(event); depth.current += 1; setDragging(true); }}
      onDragOver={stop}
      onDragLeave={(event) => {
        stop(event);
        depth.current = Math.max(0, depth.current - 1);
        if (depth.current === 0) setDragging(false);
      }}
      onDrop={handleDrop}
    >
      <Inbox size={compact ? 20 : 28} className="dz-icon" />
      <p className="dz-title">
        {compact ? 'Drop more report files here' : 'Drop device report files here'}
      </p>
      <p className="dz-hint">
        {compact ? (
          'They join the review below — nothing is saved until you save it.'
        ) : (
          <>
            The <code>.txt</code> reports the scan script writes. Drop as many as you like —
            one row per file. Nothing is saved until you have reviewed it.
          </>
        )}
      </p>
      <Button
        variant="secondary"
        size={compact ? 'sm' : undefined}
        onClick={() => inputRef.current?.click()}
        disabled={busy}
      >
        {compact ? 'Add more files' : 'Choose files'}
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
