/**
 * Which Office is on the machine, and whether the company is entitled to it.
 *
 * The scan reports Office as the product codes Windows holds — `O365BusinessRetail`,
 * `ProPlus2021Volume`, `O365HomePremRetail`. Those codes say what was installed,
 * not what was paid for on the company tenant, so the three answers here are
 * deliberately blunt: it is a business product (Authentic), it is a consumer or
 * personal product standing in for one (Unlicensed), or the scan found nothing
 * to judge (Undefined).
 */

/** Products bought through the company: Microsoft 365 Business, volume Office. */
const BUSINESS = /o365business|o365proplus|proplus|volume|standard\d{4}volume|enterprise/i;

/** Personal and free products: someone's own subscription, or a free viewer. */
const CONSUMER = /homeprem|homebusiness|homestudent|personal|free|starter|mondo/i;

/** The parts of Office that never carried a licence of their own. */
const NOT_OFFICE = /onenote|teams|onedrive|skype/i;

export const LICENSE_STATUSES = ['Authentic', 'Unlicensed', 'Undefined'];

export function officeLicense(products = []) {
  const codes = products.map((code) => String(code).trim()).filter(Boolean);
  const licensable = codes.filter((code) => !NOT_OFFICE.test(code));

  if (!licensable.length) {
    return {
      licenseStatus: 'Undefined',
      licenseNote: codes.length
        ? 'Only free Microsoft apps found — no Office licence to check'
        : 'The scan found no Microsoft Office on this machine',
      officeProducts: codes,
    };
  }

  const business = licensable.filter((code) => BUSINESS.test(code));
  if (business.length) {
    return {
      licenseStatus: 'Authentic',
      licenseNote: `Company product installed: ${business.join(', ')}`,
      officeProducts: codes,
    };
  }

  const consumer = licensable.filter((code) => CONSUMER.test(code));
  return {
    licenseStatus: 'Unlicensed',
    licenseNote: consumer.length
      ? `A personal Office is standing in for the company one: ${consumer.join(', ')}`
      : `Office is installed outside the company licence: ${licensable.join(', ')}`,
    officeProducts: codes,
  };
}
