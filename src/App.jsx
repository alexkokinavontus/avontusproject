import { useState, useEffect, useCallback } from "react";
import { fetchAllData } from "./api/azure";
import Dashboard from "./components/Dashboard";
import { LoadingScreen } from "./components/UIComponents";
import { ErrorScreen } from "./components/UIComponents";
import { ConfigModal } from "./components/UIComponents";

export default function App() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [lastRefresh, setLastRefresh] = useState(null);
  const [showConfig, setShowConfig] = useState(false);
  const [secret, setSecret] = useState(
    import.meta.env.VITE_AZURE_CLIENT_SECRET || localStorage.getItem("az_secret") || ""
  );

  const load = useCallback(async () => {
    if (!secret) {
      setShowConfig(true);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const result = await fetchAllData();
      setData(result);
      setLastRefresh(new Date());
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }, [secret]);

  useEffect(() => {
    if (secret) load();
    else setShowConfig(true);
  }, []);

  const handleSecretSave = (s) => {
    localStorage.setItem("az_secret", s);
    setSecret(s);
    setShowConfig(false);
    setTimeout(load, 100);
  };

  return (
    <div className="app">
      {showConfig && (
        <ConfigModal
          onSave={handleSecretSave}
          onClose={() => setShowConfig(false)}
          hasSecret={!!secret}
        />
      )}
      {loading && <LoadingScreen />}
      {error && !loading && (
        <ErrorScreen error={error} onRetry={load} onConfig={() => setShowConfig(true)} />
      )}
      {data && !loading && !error && (
        <Dashboard
          data={data}
          lastRefresh={lastRefresh}
          onRefresh={load}
          onConfig={() => setShowConfig(true)}
        />
      )}
      {!data && !loading && !error && !showConfig && (
        <div className="empty-state">
          <button onClick={() => setShowConfig(true)} className="btn-primary">
            Configure Azure Connection
          </button>
        </div>
      )}
    </div>
  );
}
