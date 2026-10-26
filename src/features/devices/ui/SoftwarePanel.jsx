import { useEffect, useId, useMemo, useRef, useState } from 'react';
import { Card } from '../../../components/ui/Surfaces';
import Pager from '../../../components/ui/Pager';
import { paginate } from '../../../components/ui/paginate';
import { Copy } from '../../../components/ui/Icons';
import { softwareOf } from '../derive/deriveSoftware';
import { scanKindOf } from '../licenceView';
import {
  backgroundCount, copyTextOf, highlightParts, installedLabel, kindTagOf, visibleSoftware,
} from '../softwareView';

const PAGE_SIZE = 25;

const plural = (count, word) => `${count} ${word}${count === 1 ? '' : 's'}`;

/**
 * The Installed software section: a search that filters as you type, the
 * runtimes and drivers folded away by default, Copy list, and the rows a page
 * at a time. A machine with no software on record shows no section at all.
 * The list itself is the evidence, not the scan version: a row whose version
 * was never stored still shows the programs it does hold.
 */
export default function SoftwarePanel({ device }) {
  const headingId = useId();
  const searchId = useId();
  const switchId = useId();
  const list = useMemo(() => softwareOf(device), [device]);
  const scanKind = scanKindOf(device);

  const [query, setQuery] = useState('');
  const [hideRuntimes, setHideRuntimes] = useState(true);
  const [size, setSize] = useState(PAGE_SIZE);
  // The page is remembered with the filter it was reached under, so a new search
  // starts at page 1 without an effect resetting it after the first paint.
  const [pageAt, setPageAt] = useState({ key: '', page: 1 });
  const [copied, setCopied] = useState(0);
  const timer = useRef(null);

  useEffect(() => () => clearTimeout(timer.current), []);

  const filterKey = `${device.id ?? device.computerName}|${query}|${hideRuntimes}|${size}`;
  const folded = backgroundCount(list);
  const shown = useMemo(
    () => visibleSoftware(list, { query, hideRuntimes }),
    [list, query, hideRuntimes],
  );
  const page = paginate(shown, pageAt.key === filterKey ? pageAt.page : 1, size);
  const canCopy = typeof navigator !== 'undefined' && Boolean(navigator.clipboard?.writeText);

  const copyList = async () => {
    if (!canCopy || shown.length === 0) return;
    try {
      await navigator.clipboard.writeText(copyTextOf(shown));
      setCopied(shown.length);
      clearTimeout(timer.current);
      timer.current = setTimeout(() => setCopied(0), 2000);
    } catch {
      // Nothing was copied, and the list is still on screen to select by hand.
    }
  };

  const showRuntimes = () => setHideRuntimes(false);

  if (list.length === 0) return null;

  return (
    <Card className="mh-section sw" aria-labelledby={headingId} role="region">
      <div className="mh-section-head">
        <h2 id={headingId} className="mh-section-title">
          Installed software
          <span className="mh-section-count"> · {plural(list.length, 'program')}</span>
        </h2>
        {scanKind === 'basic' && <span className="sw-badge">Basic scan</span>}
      </div>

      <>
        {scanKind === 'basic' && (
          <p className="sw-note">
            Read with cmd and the registry, because PowerShell would not run. The list is complete; a few install dates may be missing.
          </p>
        )}

        <div className="sw-tools">
          <div className="sw-search">
            <label className="sw-search-label" htmlFor={searchId}>Search installed software</label>
            <input
              id={searchId}
              type="search"
              className="sw-find"
              value={query}
              placeholder="Search by name or publisher"
              onChange={(event) => setQuery(event.target.value)}
            />
          </div>
          <div className="sw-tools-end">
            <label className="sw-switch" htmlFor={switchId}>
              <input
                id={switchId}
                type="checkbox"
                checked={hideRuntimes}
                onChange={(event) => setHideRuntimes(event.target.checked)}
              />
              Hide runtimes and drivers ({folded})
            </label>
            <button
              type="button"
              className="sw-copy"
              onClick={copyList}
              disabled={!canCopy || shown.length === 0}
            >
              {copied ? (
                <span className="sw-copied" role="status">Copied {copied} lines</span>
              ) : (
                <><Copy size={15} /> Copy list</>
              )}
            </button>
          </div>
        </div>

        {shown.length === 0 ? (
          <div className="sw-empty">
            {query.trim()
              ? <>No program called “{query.trim()}” on this machine.</>
              : 'Every program here is a runtime or a driver.'}
            {hideRuntimes && folded > 0 && (
              <>
                <br />
                Runtimes and drivers are hidden —{' '}
                <button type="button" className="sw-link" onClick={showRuntimes}>search them too</button>.
              </>
            )}
          </div>
        ) : (
          <>
            <div className="sw-table" role="table" aria-label="Installed programs">
              <div className="sw-row sw-head" role="row">
                <span role="columnheader">Program</span>
                <span role="columnheader">Version</span>
                <span role="columnheader">Publisher</span>
                <span role="columnheader">Installed</span>
              </div>
              {page.rows.map((entry, index) => {
                const tag = kindTagOf(entry);
                return (
                  <div key={`${entry.name}|${entry.version}|${index}`}
                    className={`sw-row${tag ? ' sw-row-background' : ''}`} role="row">
                    <span role="cell" className="sw-name">
                      {highlightParts(entry.name, query).map((part, at) => (part.match
                        ? <b key={at}>{part.text}</b>
                        : <span key={at}>{part.text}</span>))}
                      {tag && <span className="sw-tag">{tag}</span>}
                      <span className="sw-meta">
                        <span className="sw-version">{entry.version || '—'}</span>
                        <span className="sw-sep" aria-hidden="true"> · </span>
                        <span className="sw-publisher">{entry.publisher || '—'}</span>
                      </span>
                    </span>
                    <span role="cell" className="sw-version-col">{entry.version || '—'}</span>
                    <span role="cell" className="sw-publisher-col">{entry.publisher || '—'}</span>
                    <span role="cell" className="sw-installed">{installedLabel(entry.installedOn)}</span>
                  </div>
                );
              })}
            </div>
            <Pager
              page={page}
              size={size}
              label="programs"
              onPage={(next) => setPageAt({ key: filterKey, page: next })}
              onSize={setSize}
            />
          </>
        )}
      </>
    </Card>
  );
}
