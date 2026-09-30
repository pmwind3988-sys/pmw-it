import { useEffect } from 'react';
import AppShell from '../components/AppShell';
import { Card } from '../components/ui/Surfaces';
import { useSession } from '../hooks/useSession';

/**
 * The landing screen, emptied while it is redesigned. It carries nothing but a
 * label saying so; the request figures live on `/requests`.
 */
export default function DashboardPage() {
  const { markContentReady } = useSession();

  useEffect(() => {
    document.title = 'PMW IT — Dashboard';
  }, []);

  // The sign-in veil waits for the landing screen to say it is ready. There is
  // nothing to load here any more, so it is ready as soon as it mounts.
  useEffect(() => {
    markContentReady();
  }, [markContentReady]);

  return (
    <AppShell title="Dashboard">
      <Card className="dash-dev">
        <span className="dash-dev-label">In development</span>
        <p>This dashboard is being rebuilt. Check back soon.</p>
      </Card>
    </AppShell>
  );
}
