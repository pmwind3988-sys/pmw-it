import WorldMap from './WorldMap';
import LocationView from './LocationView';
import DepartmentView from './DepartmentView';
import { EmptyState } from '../../../components/ui/Surfaces';

/** Which level of the map the address asks for. */
export default function DeviceMap({ devices, loading, params, standard }) {
  const location = params.get('location') ?? '';
  const department = params.get('department') ?? '';
  const place = params.get('place') ?? '';

  if (loading && devices.length === 0) return <EmptyState>Loading the register…</EmptyState>;
  if (devices.length === 0) return <EmptyState>Nothing in the register yet. Open the Import tab and drop your scan reports.</EmptyState>;
  if (place) return <DepartmentView devices={devices} place={place} />;
  if (location && department) return <DepartmentView devices={devices} location={location} department={department} />;
  if (location) return <LocationView devices={devices} location={location} standard={standard} />;
  return <WorldMap devices={devices} />;
}
