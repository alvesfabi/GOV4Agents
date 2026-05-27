import { Routes, Route, Navigate } from 'react-router-dom';
import { AppShell } from '../components/AppShell';
import { IntroPage } from '../features/intro/IntroPage';
import { SetupPage } from '../features/setup/SetupPage';
import { ManagePage } from '../features/manage/ManagePage';
import { GovernPage } from '../features/govern/GovernPage';
import { ProtectPage } from '../features/protect/ProtectPage';

export function AppRoutes() {
  return (
    <Routes>
      <Route element={<AppShell />}>
        <Route index element={<IntroPage />} />
        <Route path="setup" element={<SetupPage />} />
        <Route path="setup/*" element={<Navigate to="/setup" replace />} />
        <Route path="manage" element={<ManagePage />} />
        <Route path="govern" element={<GovernPage />} />
        <Route path="protect" element={<ProtectPage />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Route>
    </Routes>
  );
}
