import Field from '../form/Field';
import {
  TextInput, TextArea, NumberInput, DateInput, SelectInput,
} from '../form/Inputs';
import { CheckList } from '../form/Choices';
import RepeatRows from '../form/RepeatRows';
import { Lock } from '../ui/Icons';
import {
  ENTITIES, CHECKLIST_ITEMS, REQUESTABLE_ITEMS, isRequest, newItemRow,
} from '../../features/forms/checklistForm';
import { FIELD_LABELS, describeValue } from '../../features/forms/describe';

/**
 * The body of the asset checklist — everything between the form type and the
 * signature — drawn once for the three places that need it: the portal's own
 * checklist, the link builder IT pre-fills, and the public page an employee
 * opens from a link.
 *
 * `locked` names fields to show as text rather than inputs (what IT fixed on a
 * link). `adornment(field)` puts something beside a field's label (the
 * builder's "employee can edit" switch). Neither decides anything: which
 * fields are locked is `linkRules.js`'s answer, passed in.
 */

function LockedValue({ field, value }) {
  const shown = describeValue(field, value);
  const lines = Array.isArray(shown) ? shown : [shown];

  return (
    <div className="ff-locked" title="Filled in by IT">
      <Lock size={13} aria-hidden="true" />
      {lines.filter(Boolean).length ? (
        <ul>{lines.map((line) => <li key={line}>{line}</li>)}</ul>
      ) : (
        <span className="ff-locked-empty">—</span>
      )}
      <span className="ff-sr">(filled in by IT, cannot be changed)</span>
    </div>
  );
}

export default function ChecklistFields({
  values, errors = {}, update, locked = [], adornment, requireDetails = true,
}) {
  const isLocked = (field) => locked.includes(field);
  const adorn = (field) => adornment?.(field) ?? null;

  const field = (name, input, { required = requireDetails, help, wide = false } = {}) => (
    <Field
      key={name}
      label={FIELD_LABELS[name]}
      htmlFor={isLocked(name) ? undefined : name}
      required={required && !isLocked(name)}
      error={errors[name]}
      help={isLocked(name) ? undefined : help}
      wide={wide}
      adornment={adorn(name)}
    >
      {isLocked(name) ? <LockedValue field={name} value={values[name]} /> : input}
    </Field>
  );

  return (
    <>
      <div className="ff-grid">
        {field('employeeName', (
          <TextInput
            id="employeeName"
            value={values.employeeName}
            onChange={update('employeeName')}
            error={errors.employeeName}
            autoComplete="name"
          />
        ))}
        {field('employeeNo', (
          <TextInput
            id="employeeNo"
            value={values.employeeNo}
            onChange={update('employeeNo')}
            error={errors.employeeNo}
          />
        ))}
        {field('position', (
          <TextInput
            id="position"
            value={values.position}
            onChange={update('position')}
            error={errors.position}
          />
        ))}
        {field('entity', (
          <SelectInput
            id="entity"
            value={values.entity}
            onChange={update('entity')}
            options={ENTITIES}
            error={errors.entity}
          />
        ))}
        {field('formDate', (
          <DateInput
            id="formDate"
            value={values.formDate}
            onChange={update('formDate')}
            error={errors.formDate}
          />
        ))}
      </div>

      {isRequest(values.formMode)
        ? field('items', (
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
        ), { required: false, help: 'Add a line for each thing needed.', wide: true })
        : field('checkedItems', (
          <CheckList
            value={values.checkedItems}
            onChange={update('checkedItems')}
            options={CHECKLIST_ITEMS}
          />
        ), { required: false, help: 'Tick everything covered by this handover.', wide: true })}

      <div className="ff-grid">
        {field('serialNumbers', (
          <TextArea
            id="serialNumbers"
            rows={2}
            value={values.serialNumbers}
            onChange={update('serialNumbers')}
          />
        ), { required: false, help: 'Optional.' })}
        {field('otherRemarks', (
          <TextArea
            id="otherRemarks"
            rows={2}
            value={values.otherRemarks}
            onChange={update('otherRemarks')}
          />
        ), { required: false, help: 'Optional.' })}
      </div>
    </>
  );
}
