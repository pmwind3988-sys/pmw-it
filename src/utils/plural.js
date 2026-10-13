/** "1 desktop", "3 desktops": a count with its noun agreeing. */
export const plural = (count, one, many = `${one}s`) => `${count} ${count === 1 ? one : many}`;
