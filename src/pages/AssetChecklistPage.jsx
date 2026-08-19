import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import { Model } from 'survey-core';
import { Survey } from 'survey-react-ui';
import 'survey-core/survey-core.min.css';
import { useMsal } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { submitAssetChecklistToSharePoint } from '../services/sharePointService';
import { sharePointRequest } from '../authConfig';
import SignatureDialog from '../components/SignatureDialog';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { RequestTypeBadge } from '../components/ui/Badges';
import { ArrowLeft } from '../components/ui/Icons';

const ASSET_ITEMS = [
  'Laptop', 'Desktop', 'Mouse', 'Monitor', 'Keyboard',
  'HDMI Cable', 'VGA Cable', 'Speaker', 'Earphone',
  'Phone & Simcard', 'Locker Key',
];

const ENTITY_CHOICES = ['pmw', 'pmw-ss', 'pmw-th'];

const getSurveyJson = (formMode) => ({
  completeText: 'Submit',
  showPreviewBeforeComplete: 'showAllQuestions',
  theme: 'default',
  elements: [
    {
      type: 'text',
      name: 'employeeName',
      title: 'Employee Name',
      isRequired: true,
      placeholder: 'Enter full name',
    },
    {
      type: 'text',
      name: 'employeeNo',
      title: 'Employee No',
      isRequired: true,
      placeholder: 'Enter employee number',
    },
    {
      type: 'text',
      name: 'position',
      title: 'Position',
      isRequired: true,
      placeholder: 'Enter position/title',
    },
    {
      type: 'dropdown',
      name: 'entity',
      title: 'Entity',
      isRequired: true,
      choices: ENTITY_CHOICES.map(v => ({ value: v, text: v })),
    },
    {
      type: 'text',
      name: 'submissionDateDisplay',
      title: 'Date & Time',
      readOnly: true,
      placeholder: 'Waiting for real-time clock...',
    },
    {
      type: 'matrixdynamic',
      name: 'assetMatrix',
      title: 'Asset Checklist',
      addRowText: 'Add Asset Item',
      removeRowText: 'Remove',
      defaultRowValue: { item: '', quantity: 1, serialNumber: '', remarks: '' },
      columns: [
        {
          name: 'item',
          title: 'Item',
          cellType: 'dropdown',
          isRequired: true,
          choices: ASSET_ITEMS.map(v => ({ value: v, text: v })),
        },
        {
          name: 'quantity',
          title: 'Quantity',
          cellType: 'text',
          inputType: 'number',
          isRequired: true,
          min: 1,
          defaultValue: 1,
        },
        {
          name: 'serialNumber',
          title: 'Serial No. (Optional)',
          cellType: 'text',
        },
        {
          name: 'remarks',
          title: 'Remarks (Optional)',
          cellType: 'comment',
        },
      ],
      rowCount: 1,
      minRowCount: 1,
      maxRowCount: 50,
      allowAddRows: true,
      allowRemoveRows: true,
    },
    {
      type: 'html',
      name: 'signatureHtml',
      title: 'Signature',
      html: '<div id="signature-placeholder" style="padding:12px 0;"><button type="button" id="signature-trigger-btn" style="padding:12px 32px;font-size:15px;font-weight:500;border:2px dashed #999;border-radius:8px;background:transparent;color:#555;cursor:pointer;width:100%;transition:all 0.2s;">Click to Sign</button></div>',
    },
  ],
});

export default function AssetChecklistPage() {
  const { instance } = useMsal();
  // Signing in, the theme toggle and signing out all belong to the shell now;
  // this page only needs the token it acquires for the submission.

  const [formMode, setFormMode] = useState(null);
  const [survey, setSurvey] = useState(null);
  const [submitState, setSubmitState] = useState('idle');
  const [toast, setToast] = useState('');
  const [formError, setFormError] = useState('');
  const [showSignature, setShowSignature] = useState(false);
  const signatureValueRef = useRef(null);

  useEffect(() => {
    document.title = 'PMW IT — Asset checklist';
  }, []);

  // Update real-time clock every second
  useEffect(() => {
    if (!survey) return;
    const updateClock = () => {
      const now = new Date();
      survey.setValue('submissionDateDisplay', now.toLocaleString('en-MY', {
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit',
        hour12: false,
      }));
    };
    updateClock();
    const interval = setInterval(updateClock, 1000);
    return () => clearInterval(interval);
  }, [survey]);

  // Handle signature button rendering inside survey
  useEffect(() => {
    if (!survey) return;

    survey.onAfterRenderQuestion.add((_, options) => {
      if (options.question.name !== 'signatureHtml') return;
      const btn = document.getElementById('signature-trigger-btn');
      if (btn) {
        btn.onclick = () => setShowSignature(true);
      }
    });

    return () => {
      survey.onAfterRenderQuestion.remove();
    };
  }, [survey]);

  const surveyModel = useMemo(() => {
    if (!formMode) return null;
    const model = new Model(getSurveyJson(formMode));
    return model;
  }, [formMode]);

  // Track survey instance for clock updates
  useEffect(() => {
    if (surveyModel) {
      setSurvey(surveyModel);
    }
  }, [surveyModel]);

  const getSharePointToken = async () => {
    const account = instance.getActiveAccount();
    if (!account) throw new Error('No signed-in account found. Please log in first.');
    try {
      const result = await instance.acquireTokenSilent({ ...sharePointRequest, account });
      return result.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const result = await instance.acquireTokenPopup({ ...sharePointRequest, account });
        return result.accessToken;
      }
      throw error;
    }
  };

  // Signature saved callback
  const handleSignatureSave = useCallback((dataUrl) => {
    signatureValueRef.current = dataUrl;
    setShowSignature(false);
    // Update the button text in the survey
    const sigBtn = document.getElementById('signature-trigger-btn');
    if (sigBtn) {
      if (dataUrl) {
        sigBtn.textContent = '✓ Signed';
        sigBtn.style.borderColor = '#10b981';
        sigBtn.style.color = '#10b981';
      } else {
        sigBtn.textContent = 'Click to Sign';
        sigBtn.style.borderColor = '#999';
        sigBtn.style.color = '#555';
      }
    }
  }, []);

  // Handle form submission
  useEffect(() => {
    if (!survey) return;

    const handleComplete = async () => {
      const data = survey.data;
      if (!data) {
        setToast('No form data to submit');
        setTimeout(() => setToast(''), 3000);
        return;
      }

      setSubmitState('submitting');
      setFormError('');

      try {
        const accessToken = await getSharePointToken();
        const now = new Date();

        const formPayload = {
          formMode,
          employeeName: data.employeeName || '',
          employeeNo: data.employeeNo || '',
          position: data.position || '',
          entity: data.entity || '',
          submissionDateISO: now.toISOString(),
          assetMatrix: data.assetMatrix || [],
          signatureDataUrl: signatureValueRef.current || null,
        };

        await submitAssetChecklistToSharePoint(
          import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk',
          accessToken,
          formPayload
        );

        setSubmitState('success');
        setToast('Form submitted successfully!');
        setTimeout(() => setToast(''), 3000);
      } catch (error) {
        console.error('[AssetChecklist] Submit error:', error);
        setSubmitState('error');
        setFormError(error.message || 'An unknown error occurred.');
      }
    };

    survey.onComplete.add(handleComplete);
    return () => {
      survey.onComplete.remove(handleComplete);
    };
  }, [survey, formMode]);

  const modeLabel =
    formMode === 'Individual Request'
      ? 'Individual request'
      : formMode === 'In'
        ? 'IN · onboarding'
        : 'OUT · offboarding';

  // ── Mode selection ──
  if (!formMode) {
    return (
      <AppShell title="Asset checklist" subtitle="Pick the handover this checklist is for.">
        <div className="mode-selector">
          <div className="mode-selector-card">
            <div className="mode-buttons">
              <button className="mode-btn mode-btn-in" onClick={() => setFormMode('In')}>
                <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4" />
                  <polyline points="10 17 15 12 10 7" />
                  <line x1="15" y1="12" x2="3" y2="12" />
                </svg>
                <span className="mode-btn-label">IN</span>
                <span className="mode-btn-desc">Onboarding employee</span>
              </button>
              <button className="mode-btn mode-btn-out" onClick={() => setFormMode('Out')}>
                <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
                  <polyline points="16 17 21 12 16 7" />
                  <line x1="21" y1="12" x2="9" y2="12" />
                </svg>
                <span className="mode-btn-label">OUT</span>
                <span className="mode-btn-desc">Offboarding employee</span>
              </button>
              <button className="mode-btn mode-btn-individual" onClick={() => setFormMode('Individual Request')}>
                <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                  <polyline points="14 2 14 8 20 8" />
                  <line x1="12" y1="18" x2="12" y2="12" />
                  <line x1="9" y1="15" x2="15" y2="15" />
                </svg>
                <span className="mode-btn-label">
                  INDIVIDUAL
                  <br />
                  REQUEST
                </span>
                <span className="mode-btn-desc">Single item request</span>
              </button>
            </div>
          </div>
        </div>
      </AppShell>
    );
  }

  // ── The checklist itself ──
  return (
    <AppShell
      title={`Asset checklist — ${modeLabel}`}
      subtitle="Confirm each item, then sign to complete the handover."
      actions={
        <>
          <RequestTypeBadge type={formMode} />
          <Button variant="ghost" icon={ArrowLeft} onClick={() => setFormMode(null)}>
            Change type
          </Button>
        </>
      }
    >
      <div className="form-content">

        {submitState === 'success' ? (
          <div className="result-card success-card">
            <div className="result-icon success-icon">
              <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2.5">
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" /><polyline points="22 4 12 14.01 9 11.01" />
              </svg>
            </div>
            <h2>Form Submitted Successfully</h2>
            <p>Your asset checklist has been saved to SharePoint.</p>
            <button className="ms-button" onClick={() => window.location.reload()}>Submit Another</button>
          </div>
        ) : submitState === 'error' ? (
          <div className="result-card error-card">
            <div className="result-icon error-icon">
              <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#ef4444" strokeWidth="2.5">
                <circle cx="12" cy="12" r="10" /><line x1="15" y1="9" x2="9" y2="15" /><line x1="9" y1="9" x2="15" y2="15" />
              </svg>
            </div>
            <h2>Submission Failed</h2>
            <p className="error-message">{formError}</p>
            <button className="ms-button" onClick={() => setSubmitState('idle')}>Try Again</button>
          </div>
        ) : submitState === 'submitting' ? (
          <div className="result-card loading-card">
            <div className="spinner"></div>
            <p>Submitting to SharePoint…</p>
          </div>
        ) : (
          <div className="survey-light-wrapper">
            {survey && <Survey model={survey} style={{ padding: '20px' }} />}
          </div>
        )}
      </div>

      {showSignature && (
        <SignatureDialog onSave={handleSignatureSave} onClose={() => setShowSignature(false)} />
      )}

      {toast && <div className="toast">{toast}</div>}
    </AppShell>
  );
}
