import { DEVICE_LIST_NAME, CHANGE_LIST_NAME, DEVICE_COLUMNS } from './deviceSchema.js';

/**
 * `LinkTitle` rather than `Title`: it renders the computer name as the link
 * into the item, which is what makes a SharePoint view navigable.
 */
const NAME = 'LinkTitle';

/**
 * The whole scan, minus the report itself. `Raw Report` is the notepad file
 * verbatim -- one cell of it is taller than the screen, and every other column
 * on the row is already a parsed piece of it.
 */
const HIDDEN_FROM_ALL_ITEMS = new Set(['RawReport']);

/**
 * Built from the schema rather than typed out, so a column added to the list
 * shows up in the default view instead of being created and never seen.
 */
const EVERY_DEVICE_FIELD = [
  NAME,
  ...DEVICE_COLUMNS
    .map((column) => column.StaticName)
    .filter((name) => !HIDDEN_FROM_ALL_ITEMS.has(name)),
];

/**
 * REST-created columns are not added to any view automatically, so a freshly
 * provisioned list shows nothing but its Title. These are the views worth
 * having: everything on the default view, and narrower cuts for the two jobs
 * the register is actually used for.
 */
export const DEVICE_VIEWS = [
  {
    list: DEVICE_LIST_NAME,
    isDefault: true,
    title: 'All Items',
    // Every field the scan produced, in schema order. Wide enough to need
    // sideways scrolling -- which is the point: nothing read out of the report
    // is hidden here, and the narrower views below are where you go to read a
    // row at a glance.
    fields: EVERY_DEVICE_FIELD,
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
