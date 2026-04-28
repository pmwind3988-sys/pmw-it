import { Routes, Route } from 'react-router-dom';
import Homepage from './pages/Homepage';
import FormPage from './pages/FormPage';
import LoginPage from './pages/LoginPage';

function App() {
  return (
    <div className="app-container">
      <Routes>
        <Route path="/"                element={<Homepage />} />
        <Route path="/login"           element={<LoginPage />} />
        <Route path="/it-boarding-form" element={<FormPage />} />
      </Routes>
    </div>
  );
}

export default App;