import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Model } from 'survey-core';
import { Survey } from 'survey-react-ui';
import 'survey-core/survey-core.min.css';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { useNavigate } from 'react-router-dom';
import { useTheme } from '../context/ThemeContext';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { ArrowLeft, Share2, Copy, Download } from '../components/ui/Icons';
import { submitEmployeesToSharePoint, fetchAllColumnChoices, fetchListItemById, updateListItem } from '../services/sharePointService';
import { sharePointRequest } from '../authConfig';
import QRCode from 'qrcode';
import { LayeredLightPanelless } from "survey-core/themes";


const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL ||
  'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const CHOICE_COLUMNS = ['Entity', 'Equipment_x0020_Items', 'Software_x0020_Licenses', 'Request_x0020_Type', 'Department'];

const getSurveyJson = (requestType, choices = {}, isEditMode = false) => {

  // Create choices with value=raw SharePoint value, text=display label
  const toChoices = (arr) =>
    arr.map(v => ({ value: v, text: v }));

  return {
    completeText: isEditMode ? 'Update' : 'Submit',
    allowAddPanel: !isEditMode,
    panelAddText: isEditMode ? null : 'Add Employee',
    panelRemoveText: isEditMode ? null : 'Remove',
    panelCount: isEditMode ? 1 : 1,
    minPanelCount: 1,
    maxPanelCount: isEditMode ? 1 : 10,
    theme: 'default',
    elements: [
      {
        type: 'paneldynamic',
        name: 'employeeRequests',
        title: isEditMode ? 'Employee Request' : 'Employee Requests',
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
               { type: 'dropdown', name: 'department', title: 'Department', isRequired: true, choices: toChoices(choices['Department'] || []) },
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
    document.title = 'PMW IT — Request form';
  }, []);

  const isAuthenticated = useIsAuthenticated();
  // The QR code is drawn to match the current theme; the toggle itself lives in
  // the shell's bar now, not on this page.
  const { isDarkMode } = useTheme();
  const navigate = useNavigate();

  const [showSharePanel, setShowSharePanel] = useState(false);
  const [qrCodeUrl, setQrCodeUrl] = useState('');
  const [toast, setToast] = useState('');
  const [formError, setFormError] = useState('');
  const [submitState, setSubmitState] = useState('idle');
  const [requestType, setRequestType] = useState('');
  const [spChoices, setSpChoices] = useState(null);
  const [choicesError, setChoicesError] = useState('');
  const [editItemId, setEditItemId] = useState(null);
  const [editItemData, setEditItemData] = useState(null);
  const sharePanelRef = useRef(null);

  // Check for edit mode from URL params
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const editId = params.get('edit');
    if (editId) {
      setEditItemId(parseInt(editId));
    }
  }, []);

  // Close share panel when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (sharePanelRef.current && !sharePanelRef.current.contains(event.target)) {
        setShowSharePanel(false);
      }
    };
    if (showSharePanel) {
      document.addEventListener('mousedown', handleClickOutside);
      return () => document.removeEventListener('mousedown', handleClickOutside);
    }
  }, [showSharePanel]);

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
        
        let itemData = null;
        if (editItemId) {
          itemData = await fetchListItemById(SHAREPOINT_SITE_URL, tokenRes.accessToken, editItemId);
        }
        
        if (!cancelled) {
          setSpChoices(choices);
          setEditItemData(itemData);
          if (itemData) {
            setRequestType(itemData.Request_x0020_Type || choices['Request_x0020_Type']?.[0] || '');
          } else {
            setRequestType(prev => prev || choices['Request_x0020_Type']?.[0] || '');
          }
        }
      } catch (err) {
        if (!cancelled) setChoicesError(err.message || 'Failed to load form options from SharePoint.');
      }
    }
    loadChoices();
    return () => { cancelled = true; };
  }, [isAuthenticated, retryCount]);

  const survey = useMemo(() => {
    if (!spChoices) return null;
    return new Model(getSurveyJson(requestType, spChoices, !!editItemId));
  }, [requestType, spChoices, editItemId]);

// Disable add/remove panels in edit mode after survey is created
  useEffect(() => {
    if (!survey || !editItemId) return;
    
    // Set paneldynamic to view-only mode via JS after survey renders
    survey.onCurrentPageChanged = () => {
      const panel = survey.getPanelByName('employeeRequests');
      if (panel) {
        panel.allowAddPanel = false;
        panel.allowRemovePanel = false;
        panel.maxPanelCount = 1;
        panel.minPanelCount = 1;
      }
    };
    
    // Also apply immediately
    const panel = survey.getPanelByName('employeeRequests');
    if (panel) {
      panel.allowAddPanel = false;
      panel.allowRemovePanel = false;
      panel.maxPanelCount = 1;
      panel.minPanelCount = 1;
    }
    
    // Direct DOM hiding after render
    const hideButtons = () => {
      setTimeout(() => {
        const allButtons = document.querySelectorAll('button');
        allButtons.forEach(el => {
          const text = (el.textContent || '').toLowerCase().trim();
          if (text === 'add employee' || text === '+ add employee') {
            el.style.display = 'none';
            el.style.visibility = 'hidden';
            el.style.height = '0';
            el.style.padding = '0';
            el.style.overflow = 'hidden';
          }
          if (text === 'remove') {
            el.style.display = 'none';
          }
        });
      }, 100);
    };
    
    hideButtons();
    const interval = setInterval(hideButtons, 500);
    setTimeout(() => clearInterval(interval), 3000);
    
    return () => clearInterval(interval);
  }, [survey, editItemId]);

  // Restore draft from localStorage or populate edit data
  useEffect(() => {
    if (!survey) return;
    
    if (editItemData && editItemId) {
      // Populate form with existing data for editing
      const employeeData = [{
        fullName: editItemData.Title || '',
        callingName: editItemData.Calling_x0020_Name || '',
        position: editItemData.Position || '',
             entity: editItemData.Entity || '',
             department: editItemData.Department || '',
             employeeId: editItemData.Employee_x0020_ID || '',
        joinDate: editItemData.Join_x0020__x002f__x0020_Last_x0 ? editItemData.Join_x0020__x002f__x0020_Last_x0.split('T')[0] : '',
        equipmentItems: editItemData.Equipment_x0020_Items ? editItemData.Equipment_x0020_Items.results : [],
        equipmentRemarks: editItemData.Equipment_x0020_Remarks || '',
        softwareLicenses: editItemData.Software_x0020_Licenses ? editItemData.Software_x0020_Licenses.results : [],
        specialPermission: editItemData.Special_x0020_Permission || '',
      }];
      survey.data = { employeeRequests: employeeData };
    } else {
      // Restore draft from localStorage for new form
      const saved = localStorage.getItem(`surveyData_${requestType}`);
      if (saved) {
        try { survey.data = JSON.parse(saved); } catch (_) { }
      }
    }
  }, [requestType, survey, editItemData, editItemId]);

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
        
        if (editItemId) {
          // Update existing item
          const emp = employees[0];
          const itemData = {
            Title: emp.fullName || '',
            Calling_x0020_Name: emp.callingName || '',
            Position: emp.position || '',
         Entity: emp.entity || '',
         Department: emp.department || '',
         Employee_x0020_ID: emp.employeeId || '',
            Equipment_x0020_Remarks: emp.equipmentRemarks || '',
            Special_x0020_Permission: emp.specialPermission || '',
          };
          if (emp.joinDate) {
            const d = new Date(emp.joinDate);
            if (!isNaN(d.getTime())) {
              itemData.Join_x0020__x002f__x0020_Last_x0 = d.toISOString();
            }
          }
          if (emp.equipmentItems?.length) itemData.Equipment_x0020_Items = emp.equipmentItems;
          if (emp.softwareLicenses?.length) itemData.Software_x0020_Licenses = emp.softwareLicenses;
          itemData.Request_x0020_Type = requestType;
          
          await updateListItem(SHAREPOINT_SITE_URL, accessToken, editItemId, itemData);
          setSubmitState('success');
          setToast('Request updated successfully!');
          setTimeout(() => setToast(''), 3000);
        } else {
          // Create new item
          await submitEmployeesToSharePoint(SHAREPOINT_SITE_URL, accessToken, employees, requestType);
          localStorage.removeItem(`surveyData_${requestType}`);
          setSubmitState('success');
          setToast('Form submitted successfully!');
          setTimeout(() => setToast(''), 3000);
        }
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
  }, [survey, requestType, editItemId]);

  // QR Code
  useEffect(() => {
    if (!showSharePanel) return;
    QRCode.toDataURL(window.location.href, {
      width: 200, margin: 2,
      color: { dark: isDarkMode ? '#FFFFFF' : '#000000', light: isDarkMode ? '#141414' : '#FFFFFF' },
    }).then(setQrCodeUrl).catch(console.error);
  }, [showSharePanel, isDarkMode]);

  const handleRetry = () => { setSubmitState('idle'); setFormError(''); };
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

  const isLoading = spChoices === null && !choicesError;

  return (
    <AppShell
      title={editItemId ? 'Request details' : 'New request'}
      subtitle={requestType ? `${requestType} request` : 'Pick a request type to begin'}
      actions={
        <>
          {editItemId && (
            <Button variant="ghost" icon={ArrowLeft} onClick={() => navigate('/requests')}>
              Back to requests
            </Button>
          )}
          {/* The request type decides which form is rendered, so it belongs with
              the page rather than in the shell's bar. It is locked while editing
              — an existing record's type is not this form's to change. */}
          <select
            value={requestType}
            onChange={(e) => setRequestType(e.target.value)}
            className="type-select"
            aria-label="Request type"
            disabled={!spChoices || !!editItemId}
          >
            {!spChoices ? (
              <option value="">Loading…</option>
            ) : (
              (spChoices?.Request_x0020_Type ?? []).map((v) => (
                <option key={v} value={v}>
                  {v}
                </option>
              ))
            )}
          </select>
          <Button variant="ghost" icon={Share2} onClick={() => setShowSharePanel((v) => !v)}>
            Share
          </Button>
        </>
      }
    >
      {showSharePanel && (
        <div className="share-panel" ref={sharePanelRef}>
          <div className="share-panel-item" onClick={handleCopyLink}>
            <Copy size={20} />
            <span>Copy link</span>
          </div>
          <div className="share-panel-item" onClick={handleDownloadQR}>
            <Download size={20} />
            <span>Download QR</span>
          </div>
          {qrCodeUrl && <img src={qrCodeUrl} alt="QR code for this form" className="share-qr-image" />}
        </div>
      )}

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
            <div className={`survey-light-wrapper ${editItemId ? 'panel-edit-mode' : ''}`}>
              <Survey model={survey} style={{ padding: '20px' }} />
            </div>
          )}
      </div>

      {toast && <div className="toast">{toast}</div>}
    </AppShell>
  );
}