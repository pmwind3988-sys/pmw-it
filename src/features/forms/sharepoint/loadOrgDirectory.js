import { spFetch, listPath, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import {
  ORG_SITE_URL, COMPANY_LIST, DEPARTMENT_LIST, companyRows, departmentRows,
} from '../orgDirectory.js';

/**
 * Reads HR's `Companies` and `Departments` lists off the HR site.
 *
 * It is a different site from the IT helpdesk one, but the token is the same:
 * the portal's scope is the ROOT domain, so whoever can open those lists in
 * SharePoint can read them here. Somebody who cannot gets an error, and the
 * checklist falls back to typing — see `AssetChecklistPage`.
 */
async function readList(token, list, select) {
  const path = `${listPath(list)}/items?$select=${select}&$top=5000`;
  const response = await spFetch(ORG_SITE_URL, path, { token, accept: ITEM_ACCEPT });
  if (!response.ok) throw new Error(`Could not read the ${list} list (${response.status})`);
  const data = await response.json();
  return data.value ?? [];
}

export async function loadOrgDirectory(token) {
  const [companies, departments] = await Promise.all([
    readList(token, COMPANY_LIST, 'Title,Code,IsActive'),
    readList(token, DEPARTMENT_LIST, 'Title,Code,Company,IsActive'),
  ]);

  return {
    companies: companyRows(companies),
    departments: departmentRows(departments),
  };
}
