import { createContext, useContext, useState, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from '../authConfig';

const AuthContext = createContext(null);

export function AuthProvider({ children }) {
  const { instance, accounts } = useMsal();
  const [user, setUser] = useState(null);
  const [isAuthenticated, setIsAuthenticated] = useState(false);

  useEffect(() => {
    if (accounts.length > 0) {
      setUser(accounts[0]);
      setIsAuthenticated(true);
    }
  }, [accounts]);

  const login = async () => {
    try {
      await instance.loginPopup(loginRequest);
    } catch (error) {
      console.error('Login failed:', error);
    }
  };

  const logout = async () => {
    try {
      await instance.logoutPopup({
        postLogoutRedirectUri: window.location.origin,
      });
      setUser(null);
      setIsAuthenticated(false);
    } catch (error) {
      console.error('Logout failed:', error);
    }
  };

  const getAccessToken = async () => {
    try {
      const response = await instance.acquireTokenSilent(loginRequest);
      return response.accessToken;
    } catch (error) {
      const response = await instance.acquireTokenPopup(loginRequest);
      return response.accessToken;
    }
  };

  return (
    <AuthContext.Provider value={{ user, isAuthenticated, login, logout, getAccessToken }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error('useAuth must be used within AuthProvider');
  }
  return context;
}