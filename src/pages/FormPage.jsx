import React, { useState, useEffect, useMemo } from 'react';
import { Model } from 'survey-core';
import { Survey } from 'survey-react-ui';
import 'survey-core/survey-core.min.css';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { useTheme } from '../context/ThemeContext';
import { submitEmployeesToSharePoint, fetchAllColumnChoices } from '../services/sharePointService';
import { sharePointRequest } from '../authConfig';
import QRCode from 'qrcode';
import { LayeredLightPanelless } from "survey-core/themes";


const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL ||
  'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const CHOICE_COLUMNS = ['Entity', 'Equipment_x0020_Items', 'Software_x0020_Licenses', 'Request_x0020_Type'];

const getSurveyJson = (requestType, choices = {}) => {

  // Create choices with value=raw SharePoint value, text=display label
  const toChoices = (arr) =>
    arr.map(v => ({ value: v, text: v }));

  return {
    completeText: 'Submit',
    theme: 'default',
    elements: [
      {
        type: 'paneldynamic',
        name: 'employeeRequests',
        title: 'Employee Requests',
        templateElements: [
          {
            type: 'panel',
            name: 'personalInfo',
            title: 'Personal Information',
            colCount: 2,  // ← this works inside paneldynamic
            elements: [
              { type: 'text', name: 'fullName', title: 'Full Name (As per IC)', isRequired: true, placeholder: 'Enter full name' },
              { type: 'text', name: 'callingName', title: 'Calling Name', placeholder: 'Nickname (optional)' },
              { type: 'text', name: 'position', title: 'Position/Title', isRequired: true, placeholder: 'Enter position' },
              { type: 'dropdown', name: 'entity', title: 'Entity', isRequired: true, choices: toChoices(choices['Entity'] || []) },
              { type: 'text', name: 'employeeId', title: 'Employee ID', placeholder: 'Enter employee ID (optional)' },
              { type: 'text', name: 'joinDate', title: requestType?.toLowerCase() === 'onboarding' ? 'Join Date' : 'Last Working Date', isRequired: true, inputType: 'date', defaultValueExpression: 'today()' },
            ],
          },
          {
            type: 'panel',
            name: 'equipmentInfo',
            title: 'Equipment Needs',
            elements: [
              {
                type: 'checkbox', name: 'equipmentItems', title: 'Select Equipment',
                choices: toChoices(choices['Equipment_x0020_Items'] || []),
              },
              { type: 'textarea', name: 'equipmentRemarks', title: 'Special Equipment Remarks', placeholder: 'Describe any special equipment requests...' },
            ],
          },
          {
            type: 'panel',
            name: 'softwareInfo',
            title: 'Software & Access',
            elements: [
              {
                type: 'checkbox', name: 'softwareLicenses', title: 'Software Licenses Required',
                choices: toChoices(choices['Software_x0020_Licenses'] || []),
              },
              { type: 'textarea', name: 'specialPermission', title: 'Special Permission Requests', placeholder: 'Describe any special access or permissions needed...' },
            ],
          },
        ],
        panelCount: 1,
        minPanelCount: 1,
        maxPanelCount: 10,
        templateTitle: 'Employee #{panelIndex}',
        panelAddText: 'Add Employee',
        panelRemoveText: 'Remove',
      },
      {
        type: 'html',
        name: 'reviewInfo',
        html: '<div style="text-align:center;padding:30px;background:#f5f5f5;border-radius:12px;"><p style="font-size:16px;margin-bottom:16px;">Please review all employee requests before submitting.</p><p style="color:#666;">Click Submit to send your request.</p></div>',
      },
    ],
  };
};

export default function FormPage() {
  const { instance } = useMsal();
  const [retryCount, setRetryCount] = useState(0);

  useEffect(() => {
    document.title = 'IT ONBOARDING FORM';
  }, []);

  const isAuthenticated = useIsAuthenticated();
  const { isDarkMode, toggleTheme } = useTheme();

  const [showSharePanel, setShowSharePanel] = useState(false);
  const [qrCodeUrl, setQrCodeUrl] = useState('');
  const [toast, setToast] = useState('');
  const [formError, setFormError] = useState('');
  const [submitState, setSubmitState] = useState('idle');
  const [requestType, setRequestType] = useState('');
  const [spChoices, setSpChoices] = useState(null); // null = not yet loaded
  const [choicesError, setChoicesError] = useState('');

  // Fetch all choices from SharePoint before rendering form
  useEffect(() => {
    if (!isAuthenticated) return;
    let cancelled = false;

    async function loadChoices() {
      setSpChoices(null);
      setChoicesError('');
      try {
        const account = instance.getActiveAccount();
        let tokenRes;
        try {
          tokenRes = await instance.acquireTokenSilent({ ...sharePointRequest, account });
        } catch (e) {
          if (e instanceof InteractionRequiredAuthError) {
            tokenRes = await instance.acquireTokenPopup({ ...sharePointRequest, account });
          } else throw e;
        }
        const choices = await fetchAllColumnChoices(SHAREPOINT_SITE_URL, tokenRes.accessToken, CHOICE_COLUMNS);
        console.log('[SP] choices loaded:', JSON.stringify(choices, null, 2));
        if (!cancelled) {
          setSpChoices(choices);
          setRequestType(prev => prev || choices['Request_x0020_Type']?.[0] || '');
        }
        // ← second setSpChoices removed
      } catch (err) {
        if (!cancelled) setChoicesError(err.message || 'Failed to load form options from SharePoint.');
      }
    }
    loadChoices();
    return () => { cancelled = true; };
  }, [isAuthenticated, retryCount]);

  const survey = useMemo(() => {
    if (!spChoices) return null;
    return new Model(getSurveyJson(requestType, spChoices));
  }, [requestType, spChoices]);

  // Restore draft from localStorage
  useEffect(() => {
    if (!survey) return;
    const saved = localStorage.getItem(`surveyData_${requestType}`);
    if (saved) {
      try { survey.data = JSON.parse(saved); } catch (_) { }
    }
  }, [requestType, survey]);

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

  // Autosave + submit
  useEffect(() => {
    if (!survey) return;

    const handleValueChanged = () => {
      localStorage.setItem(`surveyData_${requestType}`, JSON.stringify(survey.data));
    };

    const handleComplete = async () => {
      const employees = survey.data?.employeeRequests || [];
      if (employees.length === 0) {
        setToast('No employee data to submit');
        setTimeout(() => setToast(''), 3000);
        return;
      }
      setSubmitState('submitting');
      setFormError('');
      try {
        const accessToken = await getSharePointToken();
        await submitEmployeesToSharePoint(SHAREPOINT_SITE_URL, accessToken, employees, requestType);
        localStorage.removeItem(`surveyData_${requestType}`);
        setSubmitState('success');
        setToast('Form submitted successfully!');
        setTimeout(() => setToast(''), 3000);
      } catch (error) {
        console.error('[FormPage] Submit error:', error);
        setSubmitState('error');
        setFormError(error.message || 'An unknown error occurred.');
      }
    };

    survey.onValueChanged.add(handleValueChanged);
    survey.onComplete.add(handleComplete);
    return () => {
      survey.onValueChanged.remove(handleValueChanged);
      survey.onComplete.remove(handleComplete);
    };
  }, [survey, requestType]);

  // QR Code
  useEffect(() => {
    if (!showSharePanel) return;
    QRCode.toDataURL(window.location.href, {
      width: 200, margin: 2,
      color: { dark: isDarkMode ? '#FFFFFF' : '#000000', light: isDarkMode ? '#141414' : '#FFFFFF' },
    }).then(setQrCodeUrl).catch(console.error);
  }, [showSharePanel, isDarkMode]);

  const handleRetry = () => { setSubmitState('idle'); setFormError(''); };
  const logout = () => instance.logoutRedirect({ postLogoutRedirectUri: import.meta.env.VITE_REDIRECT_URI || 'http://localhost:5173' });
  const getInitials = (name) => {
    if (!name) return 'U';
    const parts = name.split(' ');
    return parts.length >= 2 ? (parts[0][0] + parts[1][0]).toUpperCase() : name.substring(0, 2).toUpperCase();
  };
  const handleCopyLink = async () => {
    try { await navigator.clipboard.writeText(window.location.href); setToast('Link copied to clipboard!'); }
    catch (_) { setToast('Copy failed'); }
    setTimeout(() => setToast(''), 3000);
  };
  const handleDownloadQR = () => {
    if (!qrCodeUrl) return;
    const link = document.createElement('a');
    link.download = 'it-request-form-qr.png';
    link.href = qrCodeUrl;
    link.click();
    setToast('QR code downloaded!');
    setTimeout(() => setToast(''), 3000);
  };

  if (!isAuthenticated) {
    return <div style={{ textAlign: 'center', padding: '50px' }}><p>Please log in to access this page.</p></div>;
  }

  const account = instance.getActiveAccount();

  // Loading state
  const isLoading = spChoices === null && !choicesError;

  return (
    <div className="form-page">
      {/* Top banner */}
      <div className="auth-banner">
        <div className="auth-banner-left">
          {account && (
            <>
              <div className="user-avatar">{getInitials(account.name)}</div>
              <span className="user-name">{account.name}</span>
            </>
          )}
        </div>
        <div className="auth-banner-right">
          <select value={requestType} onChange={(e) => setRequestType(e.target.value)} className="type-select" disabled={!spChoices}>
            {!spChoices ? (
              <option value=''>Loading...</option>
            ) : (
              (spChoices?.Request_x0020_Type ?? []).map(v => (
                <option key={v} value={v}>{v}</option>
              ))
            )}
          </select>
          <button className="icon-btn" onClick={() => setShowSharePanel((v) => !v)} title="Share">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="18" cy="5" r="3" /><circle cx="6" cy="12" r="3" /><circle cx="18" cy="19" r="3" />
              <line x1="8.59" y1="13.51" x2="15.42" y2="17.49" /><line x1="15.41" y1="6.51" x2="8.59" y2="10.49" />
            </svg>
          </button>
          <button className="icon-btn" onClick={toggleTheme} title={isDarkMode ? 'Light Mode' : 'Dark Mode'}>
            {isDarkMode ? (
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="5" /><line x1="12" y1="1" x2="12" y2="3" /><line x1="12" y1="21" x2="12" y2="23" />
                <line x1="4.22" y1="4.22" x2="5.64" y2="5.64" /><line x1="18.36" y1="18.36" x2="19.78" y2="19.78" />
                <line x1="1" y1="12" x2="3" y2="12" /><line x1="21" y1="12" x2="23" y2="12" />
                <line x1="4.22" y1="19.78" x2="5.64" y2="18.36" /><line x1="18.36" y1="5.64" x2="19.78" y2="4.22" />
              </svg>
            ) : (
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
              </svg>
            )}
          </button>
          <button className="icon-btn" onClick={logout} title="Logout">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
              <polyline points="16 17 21 12 16 7" /><line x1="21" y1="12" x2="9" y2="12" />
            </svg>
          </button>
        </div>
      </div>

      {/* Share panel */}
      {showSharePanel && (
        <div className="share-panel">
          <div className="share-panel-item" onClick={handleCopyLink}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
              <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
            </svg>
            <span>Copy Link</span>
          </div>
          <div className="share-panel-item" onClick={handleDownloadQR}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <rect x="3" y="3" width="18" height="18" rx="2" />
              <path d="M7 7h.01M7 12h.01M7 17h.01M12 7h.01M12 12h.01M12 17h.01M17 7h.01M17 12h.01M17 17h.01" />
            </svg>
            <span>Download QR</span>
          </div>
          {qrCodeUrl && <img src={qrCodeUrl} alt="QR Code" className="share-qr-image" />}
        </div>
      )}

      {/* Main form */}
      <div className="form-container">
        <div className="form-header">
          <h1>IT Request Form</h1>
          <p>{requestType ? `${requestType} Request` : ''}</p>
        </div>

        <div className="form-content">
          {isLoading ? (
            <div className="success-screen">
              <p style={{ fontSize: 16, color: '#666' }}>Loading form options from SharePoint…</p>
            </div>

          ) : choicesError ? (
            <div className="error-screen">
              <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="#ef4444" strokeWidth="2">
                <circle cx="12" cy="12" r="10" /><line x1="15" y1="9" x2="9" y2="15" /><line x1="9" y1="9" x2="15" y2="15" />
              </svg>
              <h2>Failed to Load Form</h2>
              <p className="error-message">{choicesError}</p>
              <button className="ms-button" onClick={() => { setChoicesError(''); setRetryCount(c => c + 1); }}>
                Retry
              </button>
            </div>

          ) : submitState === 'success' ? (
            <div className="result-card success-card">
              <div className="result-icon success-icon">
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2.5">
                  <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" /><polyline points="22 4 12 14.01 9 11.01" />
                </svg>
              </div>
              <h2>Form Submitted Successfully</h2>
              <p>Your request has been saved to SharePoint.</p>
              <button className="ms-button" onClick={() => window.location.reload()}>Submit Another Request</button>
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
              <button className="ms-button" onClick={handleRetry}>Try Again</button>
            </div>

          ) : submitState === 'submitting' ? (
            <div className="result-card loading-card">
              <div className="spinner"></div>
              <p>Submitting to SharePoint…</p>
            </div>

          ) : (
            <div className="survey-light-wrapper">
              <Survey model={survey} style={{ padding: '20px' }} />
            </div>
          )}
        </div>
      </div>

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}