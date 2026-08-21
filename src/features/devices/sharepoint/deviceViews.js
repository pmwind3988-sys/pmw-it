import { DEVICE_LIST_NAME, CHANGE_LIST_NAME } from './deviceSchema.js';

/**
 * `LinkTitle` rather than `Title`: it renders the computer name as the link
 * into the item, which is what makes a SharePoint view navigable.
 */
const NAME = 'LinkTitle';

/**
 * REST-created columns are not added to any view automatically, so a freshly
 * provisioned list shows nothing but its Title. These are the views worth
 * having, and the raw multi-line columns are deliberately absent from all of
 * them — one `Raw Report` cell makes a row taller than the screen.
 */
export const DEVICE_VIEWS = [
  {
    list: DEVICE_LIST_NAME,
    isDefault: true,
    title: 'All Items',
    // The same columns and order as the portal's register, so the two screens
    // cannot tell different stories about the same machine.
    fields: [
      NAME, 'Owner', 'Department', 'DeviceType', 'ComputerModel', 'CpuModel',
      'InstalledRamGB', 'StorageTotalGB', 'StorageType', 'WindowsVersion',
      'AntivirusStatus', 'RiskLevel', 'ScannedOn',
    ],
  },
  {
    list: DEVICE_LIST_NAME,
    title: 'Needs attention',
    fields: [
      NAME, 'Owner', 'Department', 'RiskLevel', 'RiskScore', 'RiskReasons',
      'WindowsVersion', 'AntivirusStatus', 'InstalledRamGB', 'ScannedOn',
    ],
    // Risk Reasons earns its place here and nowhere else: the list is already
    // narrowed to the machines being acted on.
    query:
      '<Where><Or>'
      + '<Eq><FieldRef Name="RiskLevel" /><Value Type="Text">Critical</Value></Eq>'
      + '<Eq><FieldRef Name="RiskLevel" /><Value Type="Text">High</Value></Eq>'
      + '</Or></Where>'
      + '<OrderBy><FieldRef Name="RiskScore" Ascending="FALSE" /></OrderBy>',
  },
  {
    list: DEVICE_LIST_NAME,
    title: 'Upgrade candidates',
    fields: [
      NAME, 'Owner', 'Department', 'InstalledRamGB', 'RamSlotsUsed',
      'RamSlotsTotal', 'RamType', 'ComputerModel', 'ScannedOn',
    ],
    // A free slot and 8 GB or less: fixable with one stick, not a new machine.
    query:
      '<Where><And>'
      + '<Eq><FieldRef Name="RamUpgradable" /><Value Type="Boolean">1</Value></Eq>'
      + '<Leq><FieldRef Name="InstalledRamGB" /><Value Type="Number">8</Value></Leq>'
      + '</And></Where>'
      + '<OrderBy><FieldRef Name="InstalledRamGB" Ascending="TRUE" /></OrderBy>',
  },
  {
    list: CHANGE_LIST_NAME,
    isDefault: true,
    title: 'All Items',
    fields: [
      NAME, 'FieldName', 'OldValue', 'NewValue', 'ChangeType', 'ChangedOn', 'ChangedBy',
    ],
    query: '<OrderBy><FieldRef Name="ChangedOn" Ascending="FALSE" /></OrderBy>',
  },
];
