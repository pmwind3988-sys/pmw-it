// Files for the ticket: picked with a button here, or pasted anywhere on the
// form (RaiseTicket catches the paste and hands the files over the same way).
import * as React from 'react';
import { formatBytes } from './ticketFormat';
import { ClipIcon, CloseIcon } from './icons';
import styles from './HelpdeskTicket.module.scss';

export interface AttachmentsProps {
  files: File[];
  sizeNote: string;
  onAdd: (files: File[]) => void;
  onRemove: (index: number) => void;
}

export function Attachments(props: AttachmentsProps): React.ReactElement {
  const { files, sizeNote, onAdd, onRemove } = props;
  const inputRef = React.useRef<HTMLInputElement>(null);
  return (
    <div className={styles.field}>
      <div className={styles.attachRow}>
        <button type="button" className={styles.secondaryBtn} onClick={() => inputRef.current?.click()}>
          <ClipIcon />
          Add attachments
        </button>
        <span className={styles.hint}>Screenshots help. You can also paste one. Up to 10 MB each.</span>
      </div>
      <input
        ref={inputRef}
        type="file"
        multiple
        hidden
        tabIndex={-1}
        aria-hidden="true"
        onChange={(e) => {
          const picked = Array.from(e.target.files ?? []);
          e.target.value = '';
          if (picked.length) onAdd(picked);
        }}
      />
      {sizeNote && (
        <div className={styles.inlineError} role="alert">
          {sizeNote}
        </div>
      )}
      {files.length > 0 && (
        <ul className={styles.fileList} aria-label="Attached files">
          {files.map((f, i) => (
            <li key={`${i}-${f.name}-${f.size}`} className={styles.fileItem}>
              <span className={styles.fileName}>{f.name}</span>
              <span className={styles.fileSize}>{formatBytes(f.size)}</span>
              <button type="button" className={styles.iconBtn} aria-label={`Remove ${f.name}`} onClick={() => onRemove(i)}>
                <CloseIcon size={14} />
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
