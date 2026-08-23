import { useCallback, useEffect, useRef, useState } from 'react';
import { useIsAuthenticated } from '@azure/msal-react';
import { useNavigate } from 'react-router-dom';
import QRCode from 'qrcode';
import { useTheme } from '../context/ThemeContext';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner } from '../components/ui/Surfaces';
import {
  ArrowLeft, Share2, Copy, Download, Plus, Trash2, Check, AlertTriangle,
} from '../components/ui/Icons';
import Field from '../components/form/Field';
import { TextInput, TextArea, DateInput, SelectInput } from '../components/form/Inputs';
import { CheckList } from '../components/form/Choices';
import {
  submitEmployeesToSharePoint, fetchAllColumnChoices, fetchListItemById, updateListItem,
} from '../services/sharePointService';
import { useSharePointToken } from '../hooks/useRequests';
import {
  CHOICE_COLUMNS, MAX_EMPLOYEES, newEmployee, employeeFromItem, employeeToItem, dateLabel,
} from '../features/forms/requestForm';
import { validateRequest, hasErrors } from '../features/forms/validate';
import { ENTITIES } from '../features/forms/checklistForm';
import { mergeChoices } from '../features/sharepoint/provision';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * The onboarding / offboarding request — HR or a manager raising an event.
 *
 * Not the same thing as `/asset-checklist`, which is an employee signing for
 * kit. This one feeds `IT Request Form`, which the dashboard and `/requests`
 * are built on, and it asks exactly what it always asked. Only SurveyJS is
 * gone: with it went a three-second interval that hunted the DOM for buttons to
 * hide in edit mode.
 *
 * The dropdown options are still read LIVE from the SharePoint columns, so
 * somebody adding a department there sees it here without a deploy.
 */

const DRAFT_KEY = (requestType) => `requestDraft_${requestType}`;

export default function FormPage() {
  // Every token on this page comes through the session guard, so an expiry
  // mid-form is a dialog and a silent re-sign-in rather than a dead submit.
  const getToken = useSharePointToken();
  const isAuthenticated = useIsAuthenticated();
  const { isDarkMode } = useTheme();
  const navigate = useNavigate();

  const [choices, setChoices] = useState(null);
  const [choicesError, setChoicesError] = useState('');
  const [retry, setRetry] = useState(0);
  const [requestType, setRequestType] = useState('');

  const [editId, setEditId] = useState(null);
  const [employees, setEmployees] = useState([newEmployee()]);
  const [errors, setErrors] = useState({});
  const [state, setState] = useState('idle');
  const [failure, setFailure] = useState('');
  const [toast, setToast] = useState('');

  const [sharing, setSharing] = useState(false);
  const [qrUrl, setQrUrl] = useState('');
  const shareRef = useRef(null);

  useEffect(() => {
    document.title = 'PMW IT — Request form';
  }, []);

  // Read once from the address rather than kept in state and synced: the id in
  // the URL is the only source of truth for which record is open.
  const [editParam] = useState(
    () => new URLSearchParams(window.location.search).get('edit'),
  );

  const isEdit = Boolean(editId);

  /**
   * The options and, when editing, the record. One effect because the record's
   * request type has to win over the default the options would otherwise set.
   */
  useEffect(() => {
    if (!isAuthenticated) return undefined;
    let cancelled = false;

    (async () => {
      setChoices(null);
      setChoicesError('');
      try {
        const tokenRes = await getToken();
        const loaded = await fetchAllColumnChoices(
          SHAREPOINT_SITE_URL, tokenRes.accessToken, CHOICE_COLUMNS,
        );

        let record = null;
        if (editParam) {
          record = await fetchListItemById(
            SHAREPOINT_SITE_URL, tokenRes.accessToken, Number(editParam),
          );
        }
        if (cancelled) return;

        setChoices(loaded);

        if (record) {
          setEditId(Number(editParam));
          setEmployees([employeeFromItem(record)]);
          setRequestType(record.Request_x0020_Type || loaded.Request_x0020_Type?.[0] || '');
          return;
        }

        const type = loaded.Request_x0020_Type?.[0] || '';
        setRequestType(type);

        // A draft only comes back for a NEW request. Restoring one over a
        // record somebody opened to edit would silently overwrite it.
        const saved = localStorage.getItem(DRAFT_KEY(type));
        if (saved) {
          try {
            const parsed = JSON.parse(saved);
            if (Array.isArray(parsed) && parsed.length) setEmployees(parsed);
          } catch {
            // A corrupt draft is not worth telling anybody about; the blank
            // form underneath it is a perfectly good outcome.
          }
        }
      } catch (thrown) {
        if (!cancelled) setChoicesError(thrown.message || 'Failed to load form options from SharePoint.');
      }
    })();

    return () => { cancelled = true; };
  }, [isAuthenticated, getToken, retry, editParam]);

  /** Autosaved so a closed tab does not cost somebody ten employees of typing. */
  useEffect(() => {
    if (isEdit || !requestType) return;
    localStorage.setItem(DRAFT_KEY(requestType), JSON.stringify(employees));
  }, [employees, requestType, isEdit]);

  useEffect(() => {
    if (!sharing) return undefined;
    const onClickOutside = (event) => {
      if (shareRef.current && !shareRef.current.contains(event.target)) setSharing(false);
    };
    document.addEventListener('mousedown', onClickOutside);
    return () => document.removeEventListener('mousedown', onClickOutside);
  }, [sharing]);

  useEffect(() => {
    if (!sharing) return;
    QRCode.toDataURL(window.location.href, {
      width: 200,
      margin: 2,
      color: { dark: isDarkMode ? '#FFFFFF' : '#000000', light: isDarkMode ? '#141414' : '#FFFFFF' },
    }).then(setQrUrl).catch(() => {});
  }, [sharing, isDarkMode]);

  const say = useCallback((message) => {
    setToast(message);
    setTimeout(() => setToast(''), 3000);
  }, []);

  const setEmployee = (index) => (patch) => setEmployees(
    (current) => current.map((entry, position) => (
      position === index ? { ...entry, ...patch } : entry
    )),
  );

  const submit = async () => {
    const found = validateRequest(employees);
    setErrors(found);
    if (hasErrors(found)) return;

    setState('submitting');
    setFailure('');
    try {
      const { accessToken } = await getToken();

      if (isEdit) {
        await updateListItem(
          SHAREPOINT_SITE_URL, accessToken, editId,
          employeeToItem(employees[0], requestType),
        );
        setState('success');
        say('Request updated');
        return;
      }

      await submitEmployeesToSharePoint(
        SHAREPOINT_SITE_URL, accessToken, employees, requestType,
      );
      localStorage.removeItem(DRAFT_KEY(requestType));
      setState('success');
      say('Request submitted');
    } catch (thrown) {
      // The typing stays on screen. Losing ten employees' details to a network
      // blink is the worst thing this form could do.
      setState('error');
      setFailure(thrown.message || 'An unknown error occurred.');
    }
  };

  const copyLink = async () => {
    try {
      await navigator.clipboard.writeText(window.location.href);
      say('Link copied');
    } catch {
      say('Copy failed');
    }
  };

  const downloadQr = () => {
    if (!qrUrl) return;
    const link = document.createElement('a');
    link.download = 'it-request-form-qr.png';
    link.href = qrUrl;
    link.click();
    say('QR code downloaded');
  };

  const loading = choices === null && !choicesError;

  return (
    <AppShell
      title={isEdit ? 'Request details' : 'New request'}
      subtitle={requestType ? `${requestType} request` : 'Pick a request type to begin'}
      actions={(
        <>
          {isEdit && (
            <Button variant="ghost" icon={ArrowLeft} onClick={() => navigate('/requests')}>
              Back to requests
            </Button>
          )}
          {/* The request type decides what the form asks, so it belongs with the
              page rather than in the shell's bar. Locked while editing — an
              existing record's type is not this form's to change. */}
          <select
            value={requestType}
            onChange={(event) => setRequestType(event.target.value)}
            className="type-select"
            aria-label="Request type"
            disabled={!choices || isEdit}
          >
            {!choices
              ? <option value="">Loading…</option>
              : (choices.Request_x0020_Type ?? []).map(
                (value) => <option key={value} value={value}>{value}</option>,
              )}
          </select>
          <Button variant="ghost" icon={Share2} onClick={() => setSharing((open) => !open)}>
            Share
          </Button>
        </>
      )}
    >
      {sharing && (
        <div className="share-panel" ref={shareRef}>
          <button type="button" className="share-panel-item" onClick={copyLink}>
            <Copy size={20} /> <span>Copy link</span>
          </button>
          <button type="button" className="share-panel-item" onClick={downloadQr}>
            <Download size={20} /> <span>Download QR</span>
          </button>
          {qrUrl && <img src={qrUrl} alt="QR code for this form" className="share-qr-image" />}
        </div>
      )}

      {loading && (
        <Card className="ff-progress">
          <span className="spinner" /> <span>Loading form options from SharePoint…</span>
        </Card>
      )}

      {choicesError && (
        <ErrorBanner
          message={choicesError}
          onRetry={() => { setChoicesError(''); setRetry((n) => n + 1); }}
        />
      )}

      {state === 'success' && (
        <Card className="ff-done">
          <Check size={28} />
          <h2>{isEdit ? 'Request updated' : 'Request submitted'}</h2>
          <p>It has been saved to SharePoint.</p>
          <div className="ff-done-actions">
            <Button onClick={() => navigate('/requests')}>See all requests</Button>
            {!isEdit && (
              <Button
                variant="secondary"
                onClick={() => { setEmployees([newEmployee()]); setState('idle'); }}
              >
                Submit another
              </Button>
            )}
          </div>
        </Card>
      )}

      {failure && state === 'error' && (
        <ErrorBanner message={failure} onRetry={() => { setState('idle'); setFailure(''); }} />
      )}

      {choices && state !== 'success' && (
        <>
          {employees.map((employee, index) => (
            // Employees have no id of their own and are only ever removed from
            // the end of an edit, so the index is a stable enough key.
                <Card className="ff-panel" key={index}>
              <div className="ff-panel-head">
                <h2 className="ff-panel-title">
                  {isEdit ? 'Employee request' : `Employee #${index + 1}`}
                </h2>
                {!isEdit && employees.length > 1 && (
                  <Button
                    variant="ghost"
                    size="sm"
                    icon={Trash2}
                    onClick={() => setEmployees(
                      employees.filter((unused, position) => position !== index),
                    )}
                  >
                    Remove
                  </Button>
                )}
              </div>

              <h3 className="ff-section">Personal information</h3>
              <div className="ff-grid">
                <Field
                  label="Full Name (As per IC)"
                  required
                  error={errors[index]?.fullName}
                  htmlFor={`fullName-${index}`}
                >
                  <TextInput
                    id={`fullName-${index}`}
                    value={employee.fullName}
                    onChange={(fullName) => setEmployee(index)({ fullName })}
                    error={errors[index]?.fullName}
                  />
                </Field>

                <Field label="Calling Name" help="Optional." htmlFor={`callingName-${index}`}>
                  <TextInput
                    id={`callingName-${index}`}
                    value={employee.callingName}
                    onChange={(callingName) => setEmployee(index)({ callingName })}
                  />
                </Field>

                <Field label="Position / Title" required error={errors[index]?.position} htmlFor={`position-${index}`}>
                  <TextInput
                    id={`position-${index}`}
                    value={employee.position}
                    onChange={(position) => setEmployee(index)({ position })}
                    error={errors[index]?.position}
                  />
                </Field>

                <Field label="Entity" required error={errors[index]?.entity} htmlFor={`entity-${index}`}>
                  <SelectInput
                    id={`entity-${index}`}
                    value={employee.entity}
                    onChange={(entity) => setEmployee(index)({ entity })}
                    /* Both, merged. The column itself is reconciled on submit,
                       but Entity is PICKED before that — so the new options are
                       offered here from the first load rather than only after
                       somebody has already submitted once. */
                    options={mergeChoices(choices.Entity ?? [], ENTITIES)}
                    error={errors[index]?.entity}
                  />
                </Field>

                <Field label="Department" required error={errors[index]?.department} htmlFor={`department-${index}`}>
                  <SelectInput
                    id={`department-${index}`}
                    value={employee.department}
                    onChange={(department) => setEmployee(index)({ department })}
                    options={choices.Department ?? []}
                    error={errors[index]?.department}
                  />
                </Field>

                <Field label="Employee ID" help="Optional." htmlFor={`employeeId-${index}`}>
                  <TextInput
                    id={`employeeId-${index}`}
                    value={employee.employeeId}
                    onChange={(employeeId) => setEmployee(index)({ employeeId })}
                  />
                </Field>

                <Field
                  label={dateLabel(requestType)}
                  required
                  error={errors[index]?.joinDate}
                  htmlFor={`joinDate-${index}`}
                >
                  <DateInput
                    id={`joinDate-${index}`}
                    value={employee.joinDate}
                    onChange={(joinDate) => setEmployee(index)({ joinDate })}
                    error={errors[index]?.joinDate}
                  />
                </Field>
              </div>

              <h3 className="ff-section">Equipment needs</h3>
              <Field label="Select equipment" wide>
                <CheckList
                  value={employee.equipmentItems}
                  onChange={(equipmentItems) => setEmployee(index)({ equipmentItems })}
                  options={choices.Equipment_x0020_Items ?? []}
                />
              </Field>
              <Field label="Special equipment remarks" wide htmlFor={`equipmentRemarks-${index}`}>
                <TextArea
                  id={`equipmentRemarks-${index}`}
                  value={employee.equipmentRemarks}
                  onChange={(equipmentRemarks) => setEmployee(index)({ equipmentRemarks })}
                  placeholder="Describe any special equipment requests…"
                />
              </Field>

              <h3 className="ff-section">Software &amp; access</h3>
              <Field label="Software licences required" wide>
                <CheckList
                  value={employee.softwareLicenses}
                  onChange={(softwareLicenses) => setEmployee(index)({ softwareLicenses })}
                  options={choices.Software_x0020_Licenses ?? []}
                />
              </Field>
              <Field label="Special permission requests" wide htmlFor={`specialPermission-${index}`}>
                <TextArea
                  id={`specialPermission-${index}`}
                  value={employee.specialPermission}
                  onChange={(specialPermission) => setEmployee(index)({ specialPermission })}
                  placeholder="Describe any special access or permissions needed…"
                />
              </Field>
            </Card>
          ))}

          <div className="ff-actions">
            {!isEdit && employees.length < MAX_EMPLOYEES && (
              <Button
                variant="secondary"
                icon={Plus}
                onClick={() => setEmployees([...employees, newEmployee()])}
              >
                Add employee
              </Button>
            )}
            <Button icon={Check} onClick={submit} disabled={state === 'submitting'}>
              {state === 'submitting' ? 'Submitting…' : (isEdit ? 'Update' : 'Submit')}
            </Button>
          </div>

          {hasErrors(errors) && (
            <p className="ff-summary" role="alert">
              <AlertTriangle size={14} />
              Some answers are still needed — they are marked above.
            </p>
          )}
        </>
      )}

      {toast && <div className="toast">{toast}</div>}
    </AppShell>
  );
}
