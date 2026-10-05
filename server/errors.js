/** A write refused because the row changed since it was read (HTTP 412). */
export class ConflictError extends Error {
  constructor(message = 'The link changed while it was being submitted') {
    super(message);
    this.name = 'ConflictError';
  }
}
