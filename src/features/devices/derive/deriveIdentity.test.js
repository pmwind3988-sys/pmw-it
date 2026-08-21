import { describe, it, expect } from 'vitest';
import { deriveIdentity, parseFileName } from './deriveIdentity.js';

const emptyFields = {
  Name: [], 'Computer Name': [], 'Computer Model': [], Motherboard: [],
  'PMW Server and credentials': [],
  'Email data files found Active or Inactive account': [],
};
const withFields = (overrides) => ({ ...emptyFields, ...overrides });

describe('parseFileName', () => {
  it('splits a bracket followed by a space', () => {
    expect(parseFileName('[FINANCE] LEMON-HP_.txt'))
      .toEqual({ bracket: 'FINANCE', stem: 'LEMON-HP' });
  });

  it('splits a bracket with no space before the name', () => {
    expect(parseFileName('[QAQC FAIRUS]HPFL05_.txt'))
      .toEqual({ bracket: 'QAQC FAIRUS', stem: 'HPFL05' });
  });

  it('handles a filename with no bracket', () => {
    expect(parseFileName('ASHRAF-PC_.txt')).toEqual({ bracket: null, stem: 'ASHRAF-PC' });
  });
});

describe('deriveIdentity — department and owner from the bracket', () => {
  it('splits a bracket holding both a department and a person', () => {
    const result = deriveIdentity(withFields({ 'Computer Name': ['HPFL05'] }),
      '[QAQC FAIRUS]HPFL05_.txt');
    expect(result.department).toBe('QAQC');
    expect(result.owner).toBe('Fairus');
    expect(result.ownerSource).toBe('Filename');
  });

  it('keeps a two-word department whole', () => {
    const result = deriveIdentity(withFields({ 'Computer Name': ['PMWP001'] }),
      '[PML GUARDHOUSE] PMWP001_.txt');
    expect(result.department).toBe('PML GUARDHOUSE');
    expect(result.owner).toBe(null);
  });

  it('reads a department-only bracket', () => {
    const result = deriveIdentity(withFields({ 'Computer Name': ['AMIR-HP'] }),
      '[ENGINEERING] AMIR-HP_.txt');
    expect(result.department).toBe('ENGINEERING');
  });
});

describe('deriveIdentity — the owner chain', () => {
  it('prefers the Name field above everything', () => {
    const result = deriveIdentity(
      withFields({ Name: ['Siti Aminah'], 'PMW Server and credentials': ['server | ashraf'] }),
      '[SALES] X_.txt');
    expect(result.owner).toBe('Siti Aminah');
    expect(result.ownerSource).toBe('Name field');
  });

  it('falls back to the server credential username', () => {
    const result = deriveIdentity(
      withFields({ 'PMW Server and credentials': ['server | ashraf'] }), 'ASHRAF-PC_.txt');
    expect(result.owner).toBe('Ashraf');
    expect(result.ownerSource).toBe('Server credential');
  });

  it('falls back to the first mailbox local part, title-cased', () => {
    const result = deriveIdentity(
      withFields({
        'Email data files found Active or Inactive account': [
          'lemon.cheong@pmw-group.com.ost | C:\\Users\\user\\a.ost',
        ],
      }), 'LEMON-HP_.txt');
    expect(result.owner).toBe('Lemon Cheong');
    expect(result.ownerSource).toBe('Email');
  });

  it('prefers a mailbox over an archive', () => {
    const result = deriveIdentity(
      withFields({
        'Email data files found Active or Inactive account': [
          'old.account@pmw-industries.com.pst | C:\\a.pst',
          'jiva.ran@pmw-group.com.ost | C:\\b.ost',
        ],
      }), 'JIVA_.txt');
    expect(result.owner).toBe('Jiva Ran');
  });

  it('reports no owner when nothing in the chain yields one', () => {
    const result = deriveIdentity(emptyFields, 'PMWP001_.txt');
    expect(result.owner).toBe(null);
    expect(result.ownerSource).toBe(null);
  });
});

describe('deriveIdentity — computer name', () => {
  it('uses the field when present', () => {
    expect(deriveIdentity(withFields({ 'Computer Name': ['ASHRAF-PC'] }), 'x_.txt').computerName)
      .toBe('ASHRAF-PC');
  });

  it('falls back to the filename stem', () => {
    expect(deriveIdentity(emptyFields, '[FINANCE] EVONNE-HP_.txt').computerName)
      .toBe('EVONNE-HP');
  });
});

describe('deriveIdentity — device type', () => {
  it('reads a desktop board before anything else', () => {
    const result = deriveIdentity(withFields({
      'Computer Model': ['MS-7D99'],
      Motherboard: ['Micro-Star International Co., Ltd. | PRO B760M-A WIFI (MS-7D99)'],
    }), 'UMAIRAH-PC_.txt');
    expect(result.deviceType).toBe('Desktop');
  });

  it('does not trust a DESKTOP- computer name over the model', () => {
    const result = deriveIdentity(withFields({
      'Computer Name': ['DESKTOP-2A3ERS8'],
      'Computer Model': ['HP EliteBook Folio 9470m'],
      Motherboard: ['Hewlett-Packard | 18DF'],
    }), 'DESKTOP-2A3ERS8_.txt');
    expect(result.deviceType).toBe('Laptop');
  });

  it('reads a Dell Precision as a laptop', () => {
    const result = deriveIdentity(withFields({
      'Computer Model': ['Precision 3490'], Motherboard: ['Dell Inc. | 0JTMW8'],
    }), 'PMWL034_.txt');
    expect(result.deviceType).toBe('Laptop');
  });

  it('reads the unset DMI product string plus an ASUS board as a desktop', () => {
    const result = deriveIdentity(withFields({
      'Computer Model': ['System Product Name'],
      Motherboard: ['ASUSTeK COMPUTER INC. | PRIME H610M-K D4'],
    }), 'PMWP001_.txt');
    expect(result.deviceType).toBe('Desktop');
  });

  it('reports Unknown rather than guessing when there is no signal', () => {
    const result = deriveIdentity(emptyFields, 'CARMEN-HP_.txt');
    expect(result.deviceType).toBe('Unknown');
    expect(result.deviceTypeConfident).toBe(false);
  });
});
