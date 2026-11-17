import { useMemo, useState } from 'react';
import TillSheet from './TillSheet';
import Button from '../../../../components/ui/Button';
import { Check, ChevronLeft, Plus } from '../../../../components/ui/Icons';
import { useSharePointToken } from '../../../../hooks/useRequests';
import { SHAREPOINT_SITE_URL } from '../../useAssets';
import { addCategory } from '../../sharepoint/addCategory';
import { categoryRefusal } from '../../categories';
import { knownMakes, recentModels } from '../count';

const COMMON = ['Laptop', 'Desktop', 'Monitor', 'Phone', 'Mouse', 'Keyboard', 'Cable', 'Accessory'];
const STEPS = ['type', 'brand', 'model', 'more'];
const same = (a, b) => String(a ?? '').trim().toLowerCase() === String(b ?? '').trim().toLowerCase();

/**
 * A code the till has never seen, asked about one question at a time: what is
 * it, what brand, what model — then whether more of the same product are on
 * the bench, each with its own serial. Answering once is the point: five
 * identical laptops used to be five "New code" lines asking all three again.
 *
 * A type missing from the list can be added here. It is written to SharePoint
 * at once (`addCategory`), because the column refuses a choice it has never
 * heard of and the save would fail later for a reason nobody could read.
 *
 * Nothing else here touches SharePoint: the answers go onto the receipt.
 * Closing the sheet leaves the line asking on the receipt, as before.
 */
export default function NewItemSheet({
  code, categories, assets, drafts, onName, onMore, onSweep, onSkip,
}) {
  const getToken = useSharePointToken();
  const [step, setStep] = useState('type');
  const [category, setCategory] = useState('');
  const [make, setMake] = useState('');
  const [model, setModel] = useState('');
  const [extra, setExtra] = useState([]);
  const [adding, setAdding] = useState(false);
  const [typed, setTyped] = useState('');
  const [busy, setBusy] = useState(false);
  const [problem, setProblem] = useState('');

  const options = useMemo(() => [...categories, ...extra.filter((name) => !categories.includes(name))], [categories, extra]);
  const common = COMMON.filter((name) => options.includes(name));
  const more = options.filter((name) => !common.includes(name));

  const makes = useMemo(() => (category ? knownMakes(category, assets, drafts) : []), [category, assets, drafts]);
  const models = useMemo(() => {
    if (!category) return [];
    const all = recentModels(category, assets, drafts, 30);
    return (make.trim() ? all.filter((entry) => same(entry.manufacturer, make)) : all).slice(0, 8);
  }, [category, make, assets, drafts]);

  const pickType = (name) => { setCategory(name); setStep('brand'); };

  const addType = async () => {
    const refusal = categoryRefusal(typed, options);
    if (refusal) { setProblem(refusal); return; }
    setBusy(true);
    setProblem('');
    try {
      const tokenRes = await getToken();
      const { category: added } = await addCategory({ siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, name: typed });
      setExtra((list) => [...list, added]);
      setTyped('');
      setAdding(false);
      pickType(added);
    } catch (failure) {
      setProblem(failure.message || 'That type could not be added');
    } finally {
      setBusy(false);
    }
  };

  /** Make and model are answered; the line is named now, before "more?". */
  const finish = (chosenModel = model) => {
    const asset = assets.find((row) => row.category === category
      && same(row.manufacturer, make) && same(row.model, chosenModel)) ?? null;
    setModel(chosenModel);
    onName({ category, manufacturer: make.trim(), model: String(chosenModel).trim(), asset });
    setStep('more');
  };

  const back = () => setStep(STEPS[Math.max(0, STEPS.indexOf(step) - 1)]);
  const at = STEPS.indexOf(step) + 1;

  return (
    <TillSheet title="New item" onClose={onSkip}>
      <p className="till-sheet-lede">
        <span className="till-mono">{code}</span> is new to the till. Step {at} of 4.
      </p>

      {step === 'type' && (
        <fieldset className="till-fieldset">
          <legend>What type is it?</legend>
          <div className="till-chips">
            {common.map((name) => (
              <button key={name} type="button" className="till-chip" onClick={() => pickType(name)}>{name}</button>
            ))}
            {more.length > 0 && (
              <select className="till-chip" value="" aria-label="Another type" onChange={(event) => event.target.value && pickType(event.target.value)}>
                <option value="">More…</option>
                {more.map((name) => <option key={name}>{name}</option>)}
              </select>
            )}
            {!adding && (
              <button type="button" className="till-chip" onClick={() => { setAdding(true); setProblem(''); }}>
                <Plus size={14} /> Add a new type…
              </button>
            )}
          </div>
          {adding && (
            <div className="till-newtype">
              <label className="till-field">
                <span>Name of the new type</span>
                <input
                  value={typed}
                  placeholder="e.g. Projector"
                  autoFocus
                  disabled={busy}
                  onChange={(event) => { setTyped(event.target.value); setProblem(''); }}
                  onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); addType(); } }}
                />
              </label>
              <div className="till-newtype-actions">
                <Button icon={Plus} loading={busy} onClick={addType}>Add type</Button>
                <button type="button" className="till-link till-link-small" disabled={busy} onClick={() => { setAdding(false); setProblem(''); }}>Cancel</button>
              </div>
              <p className={problem ? 'till-newtype-note till-newtype-problem' : 'till-newtype-note'}>
                {problem || 'It becomes a type for everybody, in the register and on handovers. New types are counted by quantity; each serial is still kept.'}
              </p>
            </div>
          )}
          {onSweep && (
            <button type="button" className="till-link till-link-small" onClick={onSweep}>
              Not sure? Sweep the box — read what it is off the packaging
            </button>
          )}
        </fieldset>
      )}

      {step === 'brand' && (
        <form className="till-fieldset" onSubmit={(event) => { event.preventDefault(); setStep('model'); }}>
          <label className="till-field">
            <span>{category} · what brand?</span>
            <input value={make} onChange={(event) => setMake(event.target.value)} placeholder="e.g. Dell" autoComplete="off" autoFocus />
          </label>
          {makes.length > 0 && (
            <div className="till-chips">
              {makes.map((name) => (
                <button key={name} type="button" className="till-chip" onClick={() => { setMake(name); setStep('model'); }}>{name}</button>
              ))}
            </div>
          )}
          <Button type="submit" icon={Check} className="till-cta">{make.trim() ? 'Next' : 'Skip — brand not known'}</Button>
        </form>
      )}

      {step === 'model' && (
        <form className="till-fieldset" onSubmit={(event) => { event.preventDefault(); if (model.trim()) finish(); }}>
          <label className="till-field">
            <span>{[make.trim(), category].filter(Boolean).join(' ')} · what model?</span>
            <input value={model} onChange={(event) => setModel(event.target.value)} placeholder="e.g. Latitude 5450" autoComplete="off" autoFocus />
          </label>
          {models.length > 0 && (
            <div className="till-chips">
              {models.map((entry) => (
                <button
                  key={`${entry.manufacturer}|${entry.model}`}
                  type="button"
                  className="till-chip"
                  onClick={() => { if (!make.trim()) setMake(entry.manufacturer); finish(entry.model); }}
                >
                  {make.trim() ? entry.model : [entry.manufacturer, entry.model].filter(Boolean).join(' ')}
                </button>
              ))}
            </div>
          )}
          <Button type="submit" icon={Check} className="till-cta" disabled={!model.trim()}>
            {model.trim() ? 'Next' : 'Type the model'}
          </Button>
        </form>
      )}

      {step === 'more' && (
        <>
          <p className="till-newitem-name">{[make.trim(), model.trim()].filter(Boolean).join(' ')} <span>· {category}</span></p>
          <p className="till-sheet-lede">Is there more of this same product, each with its own serial?</p>
          <div className="till-options">
            <button type="button" className="till-option" onClick={onMore}>
              <span className="till-option-n till-mono">1</span>
              <span className="till-option-text">
                <strong>Yes — scan the others</strong>
                <span>Scan each one’s serial. They all go in as this product.</span>
              </span>
            </button>
            <button type="button" className="till-option" onClick={onSkip}>
              <span className="till-option-n till-mono">2</span>
              <span className="till-option-text">
                <strong>No, just this one</strong>
                <span>Back to scanning.</span>
              </span>
            </button>
          </div>
        </>
      )}

      {step !== 'type' && step !== 'more' && (
        <button type="button" className="till-link till-link-small" onClick={back}>
          <ChevronLeft size={14} /> Back
        </button>
      )}
    </TillSheet>
  );
}
