import { createFakeGraph } from './fakeGraph.js';
import { createLinkApi } from './checklistLinkApi.js';
import { toLinkItem, fromLinkItem } from '../src/features/forms/links/linkSchema.js';
import { reopenFields } from '../src/features/forms/links/linkChanges.js';
import { IN, OUT, INDIVIDUAL } from '../src/features/forms/checklistForm.js';

/**
 * Three demo links for `npm run dev` with `CHECKLIST_FAKE=1`, held in memory
 * and forgotten on restart. Lets the public page be opened, filled, signed and
 * printed on a machine without the app secret:
 *
 *   /c/DemoInLink   IN, name and items locked, serials open to the employee
 *   /c/DemoOutLnk   OUT, almost everything left for the employee
 *   /c/DemoReqLnk   INDIVIDUAL REQUEST, items left for the employee
 *   /c/DemoReopen   signed, then REOPENED by IT: the signature can be kept
 *   /c/DemoEdited   signed, then EDITED by IT: the copy says so
 */

// A small drawn stroke, so a kept signature is visible on the demo copy.
const DEMO_SIGNATURE = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHgAAAAoCAYAAAA16j4lAAABXElEQVR4nO2aSQ6EMAwEeQT//ypzQkKIxSReuntcEjfkuFMaOQMsWyPNUt1AE0sLFqcFi9OCxWnB4rRgceAFr+t6eSkRmQ9W8J1YJckZGSEFW4Kzi87KByfYGo5VslWgl2gowV/DsEke6Xc2I4zgmRAMkmdE0Qv2EIQu2aO3kRrlgj3FoEr27OlrrVLBEULQBEfls9aEEMxSd7QP716+1C0THC2hWnL0uIAWnLH51fM4M+MT6YIzN75KcOa6b2uVCe71ctZLFaz+i6oaCxCCEeYhwkzMJkVw9YEno4fqfHeEC0aQe+6Fpa4HaYJRqH6ylE2oYNTg6s++jzwK9ni7g8g/vL3aeRU8EoAh/EyPDPl2TIKtQY73M4Qf6Zcp32aZwZZNON/DEn6b+EaKBdMh62oTni42lPOZT9Fqwa9QzPj5b5JC6DeUMpZ/k9XE0oLFacHitGBxWrA4LVicHztqRSm4JVLLAAAAAElFTkSuQmCC';

const DAY = 86400000;

// A stand-in for the snapshot of HR's lists a real link carries.
const OPTIONS = {
  entities: [
    { value: 'PMW', label: 'PMW Industries' },
    { value: 'PCI', label: 'PCI Engineering' },
  ],
  departments: {
    PMW: [{ value: 'ENG', label: 'Engineering' }, { value: 'IT', label: 'Information Technology' }],
    PCI: [{ value: 'QA', label: 'Quality Assurance' }, { value: 'LOG', label: 'Logistics' }],
  },
};

export async function createDevApi() {
  const expiresOn = new Date(Date.now() + 14 * DAY).toISOString();
  const link = (code, formMode, preset, editable = []) => toLinkItem(
    { code, formMode, preset, editable, expiresOn, options: OPTIONS },
    { createdByName: 'IT Support (demo)', createdByEmail: 'demo@example.test' },
  );

  const graph = createFakeGraph({
    links: [
      link('DemoInLink', IN, {
        employeeName: 'Amir Hakim',
        employeeNo: 'E-1042',
        entity: 'PMW',
        department: 'ENG',
        formDate: new Date().toISOString().slice(0, 10),
        checkedItems: ['Laptop', 'Mouse', 'Monitor'],
        serialNumbers: 'Laptop 5CG1234XYZ\nMonitor CN-0F8',
      }, ['serialNumbers']),
      link('DemoOutLnk', OUT, { entity: 'PCI' }),
      link('DemoReqLnk', INDIVIDUAL, {
        employeeName: 'Siti Aminah',
        employeeNo: 'E-2210',
        position: 'Planner',
        entity: 'PCI',
      }),
      link('DemoReopen', IN, { employeeName: 'Lee Wei', employeeNo: 'E-3301', entity: 'PMW', checkedItems: ['Laptop'] }),
      link('DemoEdited', OUT, { employeeName: 'Nur Aisyah', employeeNo: 'E-4120', entity: 'PCI', checkedItems: ['Laptop', 'Mouse'] }),
    ],
  });

  const api = createLinkApi({ graph });

  // Sign the last two, then put them in the state IT's actions would leave.
  const signAs = (code, values) => api.submit(code, {
    values: { position: 'Analyst', formDate: new Date().toISOString().slice(0, 10), signature: DEMO_SIGNATURE, ...values },
  });
  await signAs('DemoReopen', { department: 'IT', serialNumbers: 'Laptop 5CG9988' });
  await signAs('DemoEdited', { department: 'QA' });

  const rowFor = async (code) => graph.findLink(code);
  const reopen = await rowFor('DemoReopen');
  await graph.updateLink(reopen.id, reopenFields(fromLinkItem({ ...reopen.fields, id: reopen.id }), Date.now() + 14 * DAY));
  const edited = await rowFor('DemoEdited');
  await graph.updateLink(edited.id, { EditedBy: 'IT Support (demo)', EditedOn: new Date().toISOString() });

  return api;
}
