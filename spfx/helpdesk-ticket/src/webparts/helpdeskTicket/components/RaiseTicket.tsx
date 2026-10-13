// The "Raise a ticket" form, its submit steps and the done screen.
import * as React from 'react';
import { TicketsApi } from '../sharepoint/ticketsApi';
import { AREAS, areaById, categoryValue, DEFAULT_PRIORITY, effectiveKind, KIND_LABELS, Kind, PRIORITIES, Priority } from '../tickets/categories';
import { Draft, emptyDraft, missingOf, Person, toTicketItem, TITLE_MAX } from '../tickets/draft';
import { Attachments } from './Attachments';
import { CcPicker } from './CcPicker';
import { useInstanceId, useMountedRef } from './hooks';
import { CheckIcon, LockIcon, TicketIcon } from './icons';
import { imagesFromClipboard, splitBySize, tooBigMessage } from './ticketFormat';
import styles from './HelpdeskTicket.module.scss';

export interface RaiseTicketProps {
  api: TicketsApi;
  user: { name: string; email: string; id: number };
  department: string;
  onShowMine: () => void;
}

type Phase = 'form' | 'working' | 'done';

interface Submitted {
  number: string;
  subject: string;
  category: string;
  priority: Priority;
  failedFiles: string[];
}

const POLL_TRIES = 5;
const POLL_GAP_MS = 2000;

const wait = (ms: number): Promise<void> => new Promise((resolve) => window.setTimeout(resolve, ms));

const initialsOf = (name: string): string => {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return '';
  return parts.length === 1 ? parts[0].charAt(0) : parts[0].charAt(0) + parts[parts.length - 1].charAt(0);
};

const messageOf = (e: unknown): string => (e instanceof Error ? e.message : 'Something went wrong.');

export function RaiseTicket(props: RaiseTicketProps): React.ReactElement {
  const { api, user, department, onShowMine } = props;
  const [draft, setDraft] = React.useState<Draft>(emptyDraft);
  const [files, setFiles] = React.useState<File[]>([]);
  const [sizeNote, setSizeNote] = React.useState('');
  const [tried, setTried] = React.useState(false);
  const [phase, setPhase] = React.useState<Phase>('form');
  const [error, setError] = React.useState('');
  const [submitted, setSubmitted] = React.useState<Submitted | null>(null);
  // Once the ticket exists, a retry must never create a second one.
  const [createdId, setCreatedId] = React.useState<number | null>(null);
  const shotNo = React.useRef(0);
  const mounted = useMountedRef();
  const areaRef = React.useRef<HTMLSelectElement>(null);
  const kindRef = React.useRef<HTMLInputElement>(null);
  const titleRef = React.useRef<HTMLInputElement>(null);
  const ids = {
    title: useInstanceId('rt-title'),
    area: useInstanceId('rt-area'),
    pri: useInstanceId('rt-pri'),
    kind: useInstanceId('rt-kind'),
    subject: useInstanceId('rt-subject'),
    details: useInstanceId('rt-details'),
    help: useInstanceId('rt-details-help')
  };

  const update = (patch: Partial<Draft>): void => setDraft((d) => ({ ...d, ...patch }));

  const addPerson = (p: Person): void =>
    setDraft((d) =>
      d.cc.some((c) => c.email.toLowerCase() === p.email.toLowerCase()) ? d : { ...d, cc: [...d.cc, p] }
    );

  const removePerson = (email: string): void =>
    setDraft((d) => ({ ...d, cc: d.cc.filter((c) => c.email !== email) }));

  const addFiles = (incoming: File[]): void => {
    const { ok, tooBig } = splitBySize(incoming);
    setFiles((prev) => [...prev, ...ok]);
    setSizeNote(tooBigMessage(tooBig.map((f) => f.name)));
  };

  const onPaste = (e: React.ClipboardEvent<HTMLDivElement>): void => {
    const images = imagesFromClipboard(e.clipboardData, () => {
      shotNo.current += 1;
      return shotNo.current;
    });
    if (images.length) {
      e.preventDefault();
      addFiles(images);
    }
  };

  const missing = missingOf(draft);
  const area = areaById(draft.area);
  const kinds: Kind[] = area && area.kinds.length > 1 ? area.kinds : [];
  const chosenKind = effectiveKind(draft.area, draft.kind);
  const category = categoryValue(draft.area, draft.kind);

  const focusFirstMissing = (): void => {
    if (!area) areaRef.current?.focus();
    else if (!chosenKind) kindRef.current?.focus();
    else titleRef.current?.focus();
  };

  const submit = async (): Promise<void> => {
    if (phase === 'working') return;
    if (missing.length) {
      setTried(true);
      focusFirstMissing();
      return;
    }
    setPhase('working');
    setError('');
    try {
      let id = createdId;
      if (id === null) {
        const cc: Person[] = [];
        for (const p of draft.cc) {
          cc.push(p.id !== undefined ? p : { ...p, id: await api.ensureUser(p.email) });
        }
        id = await api.createTicket(toTicketItem({ ...draft, cc }, department));
        setCreatedId(id);
      }
      const failedFiles: string[] = [];
      for (const f of files) {
        try {
          await api.addAttachment(id, f);
        } catch {
          failedFiles.push(f.name);
        }
      }
      let number = '';
      for (let i = 0; i < POLL_TRIES && mounted.current; i++) {
        if (i > 0) await wait(POLL_GAP_MS);
        if (!mounted.current) return;
        number = await api.ticketNumber(id).catch(() => '');
        if (number) break;
      }
      if (!mounted.current) return;
      setSubmitted({
        number,
        subject: draft.title.trim(),
        category,
        priority: draft.priority,
        failedFiles
      });
      setPhase('done');
    } catch (e) {
      if (!mounted.current) return;
      setError(messageOf(e));
      setPhase('form');
    }
  };

  const raiseAnother = (): void => {
    setDraft(emptyDraft());
    setFiles([]);
    setSizeNote('');
    setTried(false);
    setError('');
    setSubmitted(null);
    setCreatedId(null);
    setPhase('form');
  };

  if (phase === 'done' && submitted) {
    const hasNumber = submitted.number !== '';
    return (
      <div className={styles.raise}>
        <section className={styles.doneCard} aria-labelledby={ids.title}>
          <div className={styles.doneHead}>
            <div className={styles.doneIcon}>
              <CheckIcon />
            </div>
            <div className={styles.doneText}>
              <h2 id={ids.title} className={styles.doneTitle}>
                {hasNumber ? `Ticket ${submitted.number} submitted` : 'Ticket submitted'}
              </h2>
              <span className={styles.muted}>
                {hasNumber ? '' : 'Your ticket number arrives by email in a moment. '}
                Replies go to {user.email}.
              </span>
            </div>
          </div>
          <ol aria-label="Progress" className={styles.steps}>
            <li className={styles.stepActive}>New</li>
            <li className={styles.step}>In progress</li>
            <li className={styles.step}>Completed</li>
          </ol>
          <dl className={styles.summary}>
            <dt>Requested by</dt>
            <dd>{department ? `${user.name}, ${department}` : user.name}</dd>
            <dt>Short description</dt>
            <dd>{submitted.subject}</dd>
            <dt>Category</dt>
            <dd>{submitted.category || '—'}</dd>
            <dt>Priority</dt>
            <dd>{submitted.priority}</dd>
          </dl>
          {submitted.failedFiles.length > 0 && (
            <div className={styles.warnBox} role="status">
              Your ticket is saved, but these files did not upload: {submitted.failedFiles.join(', ')}. Send them to IT
              by replying to the ticket email.
            </div>
          )}
          <div className={styles.doneActions}>
            <button type="button" className={styles.primaryBtn} onClick={onShowMine}>
              My tickets
            </button>
            <button type="button" className={styles.secondaryBtn} onClick={raiseAnother}>
              Raise another
            </button>
          </div>
        </section>
      </div>
    );
  }

  const working = phase === 'working';
  const initials = initialsOf(user.name);

  return (
    <div className={styles.raise}>
      <div className={styles.layout}>
        <section className={styles.formCard} aria-labelledby={ids.title} onPaste={onPaste}>
          <div className={styles.cardHead}>
            <div className={styles.headIcon}>
              <TicketIcon />
            </div>
            <div className={styles.headText}>
              <h1 id={ids.title} className={styles.pageTitle}>
                Raise a ticket
              </h1>
              <p className={styles.muted}>
                Report something broken or ask IT for something. You&apos;ll get replies at your work email.
              </p>
            </div>
          </div>

          <div className={styles.body}>
            <p className={styles.legend}>
              <span className={styles.req} aria-hidden="true">
                *
              </span>{' '}
              marks what we need before you can send it.
            </p>

            {error && (
              <div className={styles.errorBox} role="alert">
                <span>
                  Your ticket was not sent, and your answers are still here. {error}
                </span>
                <button type="button" className={styles.secondaryBtn} onClick={() => submit()}>
                  Try again
                </button>
              </div>
            )}

            <div className={styles.twoCol}>
              <div className={styles.field}>
                <span className={styles.label}>Requested by</span>
                <div className={styles.readOnly}>
                  <span className={styles.avatar} aria-hidden="true">
                    {initials}
                  </span>
                  <span className={styles.readOnlyText}>{user.name}</span>
                  <LockIcon className={styles.lock} label="Set from your sign-in" />
                </div>
              </div>
              <div className={styles.field}>
                <span className={styles.label}>Department</span>
                <div className={styles.readOnly}>
                  <span className={styles.readOnlyText}>{department || '—'}</span>
                </div>
              </div>
            </div>
            <div className={styles.caption}>Filled in from your Microsoft 365 sign-in.</div>

            <div className={styles.twoCol}>
              <div className={styles.field}>
                <label htmlFor={ids.area} className={styles.label}>
                  <span className={styles.req} aria-hidden="true">
                    *{' '}
                  </span>
                  What&apos;s it about?
                </label>
                <select
                  id={ids.area}
                  ref={areaRef}
                  className={styles.control}
                  aria-required="true"
                  aria-invalid={tried && !area ? true : undefined}
                  value={draft.area}
                  disabled={working}
                  onChange={(e) => update({ area: e.target.value, kind: '' })}
                >
                  <option value="">Choose one</option>
                  {AREAS.map((a) => (
                    <option key={a.id} value={a.id}>
                      {a.label}
                    </option>
                  ))}
                </select>
              </div>
              <div className={styles.field}>
                <label htmlFor={ids.pri} className={styles.label}>
                  How much is it stopping you?
                </label>
                <select
                  id={ids.pri}
                  className={styles.control}
                  value={draft.priority}
                  disabled={working}
                  onChange={(e) => {
                    const p = PRIORITIES.find((x) => x.id === e.target.value);
                    update({ priority: p ? p.id : DEFAULT_PRIORITY });
                  }}
                >
                  {PRIORITIES.map((p) => (
                    <option key={p.id} value={p.id}>
                      {`${p.label} (${p.id})`}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {kinds.length > 1 && (
              <fieldset className={styles.fieldset} aria-required="true">
                <legend className={styles.label}>
                  <span className={styles.req} aria-hidden="true">
                    *{' '}
                  </span>
                  Is something broken, or do you need something?
                </legend>
                <div className={styles.radios}>
                  {kinds.map((k, i) => (
                    <label key={k} className={styles.radio}>
                      <input
                        ref={i === 0 ? kindRef : undefined}
                        type="radio"
                        name={ids.kind}
                        className={styles.radioInput}
                        checked={chosenKind === k}
                        disabled={working}
                        onChange={() => update({ kind: k })}
                      />
                      {KIND_LABELS[k]}
                    </label>
                  ))}
                </div>
              </fieldset>
            )}

            <div className={styles.field}>
              <div className={styles.labelRow}>
                <label htmlFor={ids.subject} className={styles.label}>
                  <span className={styles.req} aria-hidden="true">
                    *{' '}
                  </span>
                  Short description
                </label>
                <span className={styles.counter}>
                  {draft.title.length} / {TITLE_MAX}
                </span>
              </div>
              <input
                id={ids.subject}
                ref={titleRef}
                type="text"
                className={styles.control}
                maxLength={TITLE_MAX}
                aria-required="true"
                aria-invalid={tried && !draft.title.trim() ? true : undefined}
                placeholder="e.g. Outlook keeps asking for my password"
                value={draft.title}
                disabled={working}
                onChange={(e) => update({ title: e.target.value })}
              />
            </div>

            <div className={styles.field}>
              <label htmlFor={ids.details} className={styles.label}>
                Details
              </label>
              <span id={ids.help} className={styles.hint}>
                What were you doing, what did you see, and since when?
              </span>
              <textarea
                id={ids.details}
                rows={5}
                aria-describedby={ids.help}
                className={styles.textarea}
                value={draft.details}
                disabled={working}
                onChange={(e) => update({ details: e.target.value })}
              />
            </div>

            <CcPicker api={api} people={draft.cc} onAdd={addPerson} onRemove={removePerson} />

            <div className={styles.attachBlock}>
              <Attachments files={files} sizeNote={sizeNote} onAdd={addFiles} onRemove={(i) => setFiles((p) => p.filter((_, j) => j !== i))} />
            </div>
          </div>
        </section>

        <aside className={styles.aside} aria-label="Send">
          <div className={styles.sendCard}>
            <button type="button" className={styles.submit} onClick={() => submit()} disabled={working} aria-busy={working || undefined}>
              {working && <span className={styles.spinner} aria-hidden="true" />}
              Submit
            </button>
            {missing.length > 0 && (
              <div className={tried ? styles.missingTried : styles.missing}>
                <div className={styles.missingTitle}>Still needed</div>
                <ul className={styles.missingList}>
                  {missing.map((m) => (
                    <li key={m}>{m}</li>
                  ))}
                </ul>
              </div>
            )}
            <dl className={styles.facts}>
              <dt>Files as</dt>
              <dd>{category || '—'}</dd>
              <dt>Priority</dt>
              <dd>{draft.priority}</dd>
            </dl>
            <p className={styles.smallMuted}>IT may change the priority once they&apos;ve looked.</p>
          </div>
          <button type="button" className={styles.navCard} onClick={onShowMine}>
            My tickets <span aria-hidden="true">›</span>
          </button>
        </aside>
      </div>
    </div>
  );
}
