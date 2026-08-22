import { Suspense, lazy } from 'react';
import { Routes, Route, Navigate, useLocation } from 'react-router-dom';
import Homepage from './pages/Homepage';
import DashboardPage from './pages/DashboardPage';
import FormPage from './pages/FormPage';
import AssetChecklistPage from './pages/AssetChecklistPage';
import LoginPage from './pages/LoginPage';
import ListPage from './pages/ListPage';
import DevicesPage from './pages/DevicesPage';
import DeviceDetailPage from './pages/DeviceDetailPage';

// Lazy on purpose, not for tidiness: this page's chunk carries SheetJS and
// ECharts, which together are larger than the rest of the app. Loading it
// eagerly would make every other screen pay for a section most visits
// never open (spec section 6.3).
const DataStudioPage = lazy(() => import('./pages/DataStudioPage'));

/** `/list` was the records screen before it moved; links to it are still around. */
function LegacyListRedirect() {
  const { search } = useLocation();
  return <Navigate to={{ pathname: '/requests', search }} replace />;
}

function App() {
  return (
    <div className="app-container">
      <Routes>
        <Route path="/" element={<Homepage />} />
        <Route path="/login" element={<LoginPage />} />
        <Route path="/dashboard" element={<DashboardPage />} />
        <Route path="/requests" element={<ListPage />} />
        <Route path="/list" element={<LegacyListRedirect />} />
        <Route path="/it-boarding-form" element={<FormPage />} />
        <Route path="/asset-checklist" element={<AssetChecklistPage />} />
        <Route path="/devices" element={<DevicesPage />} />
        <Route path="/devices/:id" element={<DeviceDetailPage />} />
        <Route
          path="/data-studio"
          element={(
            <Suspense fallback={<div className="ds-route-loading">Loading Data Studio...</div>}>
              <DataStudioPage />
            </Suspense>
          )}
        />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </div>
  );
}

export default App;
