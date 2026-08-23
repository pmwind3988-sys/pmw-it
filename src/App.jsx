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
import AssetsPage from './pages/AssetsPage';
import AssetScanPage from './pages/AssetScanPage';
import AssetBatchPage from './pages/AssetBatchPage';
import AssetDetailPage from './pages/AssetDetailPage';
import AssetHandoverPage from './pages/AssetHandoverPage';
import AssetPeoplePage from './pages/AssetPeoplePage';
import AssetPersonPage from './pages/AssetPersonPage';

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
        {/* `scan` and `batch` are declared above `:id` so they are not read as
            an item id. The register's ids are numbers, but the router has no
            way to know that. */}
        <Route path="/assets" element={<AssetsPage />} />
        <Route path="/assets/scan" element={<AssetScanPage />} />
        <Route path="/assets/batch/:id" element={<AssetBatchPage />} />
        <Route path="/assets/handover" element={<AssetHandoverPage />} />
        <Route path="/assets/people" element={<AssetPeoplePage />} />
        {/* The email is URL-encoded into the path; it is the identity every
            per-person question keys on, and a name would break the moment
            somebody's changed. */}
        <Route path="/assets/people/:email" element={<AssetPersonPage />} />
        <Route path="/assets/:id" element={<AssetDetailPage />} />
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
