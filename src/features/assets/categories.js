import { CATEGORIES } from './assetKinds.js';

/**
 * What kinds of thing the register knows about, and adding one more.
 *
 * `CATEGORIES` is the list this app ships with. It was never meant to be the
 * last word — a company buys something the list has no name for roughly once a
 * year, and until now the only way to record it was "Other", which loses the
 * one fact somebody wanted to keep.
 *
 * A category somebody adds becomes a real option on the SharePoint column
 * (`sharepoint/addCategory.js`), so it is there for everybody and not just for
 * the phone that typed it. It is offered back through `categoriesIn`, which
 * reads the categories the register is ACTUALLY using rather than keeping a
 * second list of its own — two lists of the same thing is how a dropdown ends
 * up disagreeing with the rows underneath it.
 *
 * One consequence worth knowing: a category added and then not used on
 * anything is still a valid option in SharePoint, but stops being offered here
 * after a reload, because nothing is using it. Adding it again costs nothing
 * and changes nothing.
 */

/** Trimmed and squeezed, the same way a model name is. */
export function cleanCategory(typed) {
  return String(typed ?? '').trim().replace(/\s+/g, ' ');
}

/** Every category on offer: the built-in ones, then whatever is in use. */
export function categoriesIn(assets = []) {
  const list = [...CATEGORIES];
  const seen = new Set(list.map((name) => name.toLowerCase()));

  for (const asset of assets) {
    const name = cleanCategory(asset?.category);
    if (!name || seen.has(name.toLowerCase())) continue;
    seen.add(name.toLowerCase());
    list.push(name);
  }

  return list;
}

/**
 * Why this cannot be added as a category, or an empty string.
 *
 * The length cap is SharePoint's: a choice is stored in a 255-character text
 * column, and something long enough to matter here is a remark rather than a
 * category. The duplicate check is case-insensitive because "Tab" and "tab"
 * would be two options that read as one, and every count split between them.
 */
export function categoryRefusal(typed, existing = CATEGORIES) {
  const name = cleanCategory(typed);

  if (!name) return 'Type a name for the new category.';
  if (name.length > 60) return 'That is too long for a category name.';
  if (existing.some((option) => option.toLowerCase() === name.toLowerCase())) {
    return `"${name}" is already on the list.`;
  }

  return '';
}
