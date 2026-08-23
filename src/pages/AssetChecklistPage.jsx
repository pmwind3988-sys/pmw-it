import { useEffect, useState } from 'react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner } from '../components/ui/Surfaces';
import { Check, AlertTriangle, Pencil } from '../components/ui/Icons';
import Field from '../components/form/Field';
import {
  TextInput, TextArea, NumberInput, DateInput, SelectInput,
} from '../components/form/Inputs';
import { RadioCards, CheckList } from '../components/form/Choices';
import RepeatRows from '../components/form/RepeatRows';
import Wizard from '../components/form/Wizard';
import SignatureDialog from '../components/SignatureDialog';
import { useSharePointToken } from '../hooks/useRequests';
import {
  FORM_MODES, ENTITIES, CHECKLIST_ITEMS, REQUESTABLE_ITEMS, CHECKLIST_STEPS,
  emptyChecklist, newItemRow, isRequest, modeLabel,
} from '../features/forms/checklistForm';
import { validateChecklist, hasErrors } from '../features/forms/validate';
import { submitChecklist } from '../features/forms/sharepoint/submitChecklist';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * The asset checklist — an employee's signed record of what they received or
 * handed back.
 *
 * Rebuilt from the IT ASSET TRACKING FORM supplied as the reference, and
 * rebuilt WITHOUT SurveyJS: the engine styled itself so the page never looked
 * like the rest of the portal, and making it behave had needed a signature
 * button injected as raw HTML and wired up through `getElementById`.
 *
 * Which fields each mode shows, and what counts as complete, live in
 * `features/forms/` and are tested there. This file only draws them.
 */

const PHASE_LABEL = {
  provisioning: 'Preparing SharePoint',
  signature: 'Saving your signature',
  saving: 'Saving the checklist',
};

export default function AssetChecklistPage() {
  const getToken = useSharePointToken();

  const [values, setValues] = useState(emptyChecklist);
  const [step, setStep] = useState(0);
  const [errors, setErrors] = useState({});
  const [signing, setSigning] = useState(false);
  const [phase, setPhase] = useState(null);
  const [failure, setFailure] = useState('');
  const [done, setDone] = useState(false);

  useEffect(() => {
    document.title = 'PMW IT — Asset checklist';
  }, []);

  const set = (field) => (value) => setValues((current) => ({ ...current, [field]: value }));

  /**
   * Errors are recomputed as soon as one is showing, so a field stops being
   * marked the moment it is fixed — rather than staying red until the next
   * time Submit is pressed.
   */
  const update = (field) => (value) => {
    set(field)(value);
    if (hasErrors(errors)) {
      setErrors(validateChecklist({ ...values, [field]: value }, { step }));
    }
  };

  const next = () => {
    const found = validateChecklist(values, { step });
    setErrors(found);
    if (hasErrors(found)) return;
    setStep(step + 1);
    setErrors({});
  };

  const submit = async () => {
    const found = validateChecklist(values);
    setErrors(found);
    if (hasErrors(found)) return;

    if (!navigator.onLine) {
      setFailure('You are offline. Your answers are still here — try again once you have a connection.');
      return;
    }

    setFailure('');
    try {
      const tokenRes = await getToken();
      await submitChecklist({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        values,
        onProgress: setPhase,
      });
      setDone(true);
    } catch (thrown) {
      // Never cleared on failure: retyping a whole checklist because the
      // network blinked is the worst thing a form can do to somebody.
      setFailure(thrown.message || 'The checklist could not be submitted');
    } finally {
      setPhase(null);
    }
  };

  const startAgain = () => {
    setValues(emptyChecklist());
    setStep(0);
    setErrors({});
    setDone(false);
    setFailure('');
  };

  if (done) {
    return (
      <AppShell title="Asset checklist" subtitle="Signed and saved">
        <Card className="ff-done">
          <Check size={28} />
          <h2>Thank you — that is recorded</h2>
          <p>
            Your {modeLabel(values.formMode)} checklist has been saved, with your
            signature against it.
          </p>
          <Button onClick={startAgain}>Fill in another</Button>
        </Card>
      </AppShell>
    );
  }

  return (
    <AppShell
      title="IT asset tracking form"
      subtitle={step === 0
        ? 'What is this checklist for?'
        : `${modeLabel(values.formMode)} — your details and what you are signing for`}
    >
      {failure && <ErrorBanner message={failure} onRetry={submit} />}

      {phase && (
        <Card className="ff-progress">
          <span className="spinner" />
          <span>{PHASE_LABEL[phase] ?? 'Working'}…</span>
        </Card>
      )}

      <Card className="ff-panel">
        <Wizard
          steps={CHECKLIST_STEPS}
          current={step}
          onBack={() => setStep(step - 1)}
          onNext={next}
          onSubmit={submit}
          submitting={Boolean(phase)}
        >
          {step === 0 && (
            <Field
              label="Form Type"
              required
              error={errors.formMode}
              help="Pick the one this checklist is for."
            >
              <RadioCards
                name="formMode"
                value={values.formMode}
                onChange={update('formMode')}
                options={FORM_MODES}
                error={errors.formMode}
              />
            </Field>
          )}

          {step === 1 && (
            <>
              <div className="ff-grid">
                <Field label="Employee Name" htmlFor="employeeName" required error={errors.employeeName}>
                  <TextInput
                    id="employeeName"
                    value={values.employeeName}
                    onChange={update('employeeName')}
                    error={errors.employeeName}
                    autoComplete="name"
                  />
                </Field>

                <Field label="Employee No" htmlFor="employeeNo" required error={errors.employeeNo}>
                  <TextInput
                    id="employeeNo"
                    value={values.employeeNo}
                    onChange={update('employeeNo')}
                    error={errors.employeeNo}
                  />
                </Field>

                <Field label="Position" htmlFor="position" required error={errors.position}>
                  <TextInput
                    id="position"
                    value={values.position}
                    onChange={update('position')}
                    error={errors.position}
                  />
                </Field>

                <Field label="Entity" htmlFor="entity" required error={errors.entity}>
                  <SelectInput
                    id="entity"
                    value={values.entity}
                    onChange={update('entity')}
                    options={ENTITIES}
                    error={errors.entity}
                  />
                </Field>

                <Field label="Date" htmlFor="formDate" required error={errors.formDate}>
                  <DateInput
                    id="formDate"
                    value={values.formDate}
                    onChange={update('formDate')}
                    error={errors.formDate}
                  />
                </Field>
              </div>

              {isRequest(values.formMode) ? (
                <Field
                  label="What are you requesting?"
                  error={errors.items}
                  help="Add a line for each thing you need."
                  wide
                >
                  <RepeatRows
                    rows={values.items}
                    onChange={update('items')}
                    newRow={newItemRow}
                    addLabel="Add another item"
                    renderRow={(row, index, setRow) => (
                      <div className="ff-itemrow">
                        <SelectInput
                          value={row.item}
                          onChange={(item) => setRow({ ...row, item })}
                          options={REQUESTABLE_ITEMS}
                          placeholder="Choose an item…"
                          aria-label={`Item ${index + 1}`}
                        />
                        <NumberInput
                          value={row.quantity}
                          onChange={(quantity) => setRow({ ...row, quantity })}
                          aria-label={`Quantity for item ${index + 1}`}
                        />
                      </div>
                    )}
                  />
                </Field>
              ) : (
                <Field
                  label="Asset Checklist"
                  help="Tick everything covered by this handover."
                  wide
                >
                  <CheckList
                    value={values.checkedItems}
                    onChange={update('checkedItems')}
                    options={CHECKLIST_ITEMS}
                  />
                </Field>
              )}

              <div className="ff-grid">
                <Field label="Serial Numbers" htmlFor="serialNumbers" help="Optional.">
                  <TextArea
                    id="serialNumbers"
                    rows={2}
                    value={values.serialNumbers}
                    onChange={update('serialNumbers')}
                  />
                </Field>

                <Field label="Other Remarks" htmlFor="otherRemarks" help="Optional.">
                  <TextArea
                    id="otherRemarks"
                    rows={2}
                    value={values.otherRemarks}
                    onChange={update('otherRemarks')}
                  />
                </Field>
              </div>

              <Field
                label="Your Signature"
                required
                error={errors.signature}
                help="Sign in the middle of the box."
                wide
              >
                {values.signature ? (
                  <div className="ff-signed">
                    <img src={values.signature} alt="Your signature" />
                    <Button variant="ghost" size="sm" icon={Pencil} onClick={() => setSigning(true)}>
                      Sign again
                    </Button>
                  </div>
                ) : (
                  <button type="button" className="ff-signbtn" onClick={() => setSigning(true)}>
                    <Pencil size={16} /> Click to sign
                  </button>
                )}
              </Field>

              {hasErrors(errors) && (
                <p className="ff-summary" role="alert">
                  <AlertTriangle size={14} />
                  Some answers are still needed — they are marked above.
                </p>
              )}
            </>
          )}
        </Wizard>
      </Card>

      {signing && (
        <SignatureDialog
          onSave={(dataUrl) => {
            update('signature')(dataUrl || null);
            setSigning(false);
          }}
          onClose={() => setSigning(false)}
        />
      )}
    </AppShell>
  );
}
