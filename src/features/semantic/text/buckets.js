// The categories a survey answer gets filed into -- spec §7.
//
// A bucket is matched by its DESCRIPTION, not its name. "SAP" as a
// three-letter string is almost pure noise to a sentence model;
// "Problems with SAP transactions, ERP modules, master data or
// postings" sits next to the answers that belong in it. That is why the
// editor puts the description first and why renaming a bucket changes
// nothing about what lands in it.
//
// `hints` are extra phrasings, averaged in with the description. They
// exist for the cases a single sentence cannot cover -- a bucket that
// legitimately spans "the VPN drops" and "the office wifi is slow".

export const UNSORTED_ID = 'unsorted';
export const UNSORTED_LABEL = 'Unsorted';

export const STARTER_BUCKETS = [
  {
    id: 'sap',
    label: 'SAP / ERP',
    description: 'Problems working in SAP or another ERP system: transactions, modules, master data, postings and system limitations.',
    hints: ['SAP transaction is slow', 'ERP master data is wrong', 'the module cannot do what we need'],
  },
  {
    id: 'consolidation',
    label: 'Data Consolidation & Reporting',
    description: 'Combining data from several files, systems or subsidiaries into one report, and preparing recurring reports or dashboards.',
    hints: ['consolidating multiple Excel files', 'preparing the monthly management report', 'building the same dashboard again'],
  },
  {
    id: 'entry',
    label: 'Manual Data Entry',
    description: 'Retyping, copying and pasting information between systems, and transcribing from paper or documents by hand.',
    hints: ['retyping numbers into another system', 'copy and paste between spreadsheets', 'keying in data from a printout'],
  },
  {
    id: 'approvals',
    label: 'Approvals & Workflow',
    description: 'Chasing sign-off, tracking the status of a request, and sending reminders to move work along.',
    hints: ['following up on approvals', 'nobody knows the current status', 'sending reminders to get sign-off'],
  },
  {
    id: 'forms',
    label: 'Forms & Paperwork',
    description: 'Paper forms, physical signatures, hardcopy documents and routing them between people for completion.',
    hints: ['the form has to be printed and signed', 'passing hardcopy around the office', 'filling in the same form twice'],
  },
  {
    id: 'retrieval',
    label: 'Information Retrieval',
    description: 'Searching for files, records, emails or historical information, and not being able to find the latest version.',
    hints: ['hunting for the right file', 'searching old emails for a record', 'nobody knows which version is current'],
  },
  {
    id: 'communication',
    label: 'Communication & Coordination',
    description: 'Information arriving through chat groups or email instead of a system, handovers between people, and updates getting missed.',
    hints: ['updates come through WhatsApp groups', 'important messages get overlooked', 'the handover loses information'],
  },
  {
    id: 'network',
    label: 'Network & Internet',
    description: 'Connectivity problems: internet speed, VPN, remote access, shared drives and the network dropping.',
    hints: ['the internet is slow', 'the VPN keeps disconnecting', 'cannot reach the shared drive from home'],
  },
  {
    id: 'itsupport',
    label: 'IT Support & Systems',
    description: 'Hardware faults, slow computers, software problems, accounts and access rights, and waiting for IT to fix something.',
    hints: ['my laptop is very slow', 'I do not have access to the system', 'waiting for IT to respond'],
  },
  {
    id: 'digitization',
    label: 'Digitization & Automation',
    description: 'Asking for a manual process to be replaced by a system, automated, or made digital end to end.',
    hints: ['this should be automated', 'we need a proper system instead of spreadsheets', 'move the whole process online'],
  },
  {
    id: 'ai',
    label: 'AI Opportunities',
    description: 'Explicit requests for artificial intelligence, machine learning or an intelligent assistant to help with the work.',
    hints: ['AI could read these documents', 'a chatbot could answer these questions', 'machine learning to predict demand'],
  },
  {
    id: 'training',
    label: 'Training & Knowledge',
    description: 'Not knowing how to do something, undocumented processes, and knowledge that lives only in one person.',
    hints: ['nobody documented the process', 'I was never trained on this', 'only one person knows how'],
  },
];

// What actually gets embedded for a bucket. The label is deliberately
// excluded unless nothing else is left -- see the note at the top.
export function bucketPromptText(bucket) {
  const parts = [];
  const description = String(bucket?.description ?? '').trim();
  if (description !== '') parts.push(description);
  for (const hint of bucket?.hints ?? []) {
    const trimmed = String(hint ?? '').trim();
    if (trimmed !== '') parts.push(trimmed);
  }
  if (parts.length === 0) parts.push(String(bucket?.label ?? '').trim());
  return parts;
}
