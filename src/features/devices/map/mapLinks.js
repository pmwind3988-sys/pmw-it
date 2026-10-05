/** One address per level of the map, so Back and a shared link land where they should. */
export function mapHref({ location, department, place } = {}) {
  const params = new URLSearchParams({ view: 'map' });
  if (place) params.set('place', place);
  else {
    if (location) params.set('location', location);
    if (location && department) params.set('department', department);
  }
  return `/devices?${params.toString()}`;
}

export function parentHref({ location, department, place } = {}) {
  if (place) return mapHref();
  if (department) return mapHref({ location });
  return mapHref();
}
