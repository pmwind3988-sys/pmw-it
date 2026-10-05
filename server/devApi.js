import { createFakeGraph } from './fakeGraph.js';
import { createLinkApi } from './checklistLinkApi.js';
import { toLinkItem } from '../src/features/forms/links/linkSchema.js';
import { IN, OUT, INDIVIDUAL } from '../src/features/forms/checklistForm.js';

/**
 * Three demo links for `npm run dev` with `CHECKLIST_FAKE=1`, held in memory
 * and forgotten on restart. Lets the public page be opened, filled, signed and
 * printed on a machine without the app secret:
 *
 *   /c/DemoInLink   IN, name and items locked, serials open to the employee
 *   /c/DemoOutLnk   OUT, almost everything left for the employee
 *   /c/DemoReqLnk   INDIVIDUAL REQUEST, items left for the employee
 */

const DAY = 86400000;

export function createDevApi() {
  const expiresOn = new Date(Date.now() + 14 * DAY).toISOString();
  const link = (code, formMode, preset, editable = []) => toLinkItem(
    { code, formMode, preset, editable, expiresOn },
    { createdByName: 'IT Support (demo)', createdByEmail: 'demo@example.test' },
  );

  const graph = createFakeGraph({
    links: [
      link('DemoInLink', IN, {
        employeeName: 'Amir Hakim',
        employeeNo: 'E-1042',
        entity: 'PMW',
        formDate: new Date().toISOString().slice(0, 10),
        checkedItems: ['Laptop', 'Mouse', 'Monitor'],
        serialNumbers: 'Laptop 5CG1234XYZ\nMonitor CN-0F8',
      }, ['serialNumbers']),
      link('DemoOutLnk', OUT, { entity: 'PCI' }),
      link('DemoReqLnk', INDIVIDUAL, {
        employeeName: 'Siti Aminah',
        employeeNo: 'E-2210',
        position: 'Planner',
        entity: 'PML',
      }),
    ],
  });

  return createLinkApi({ graph });
}
