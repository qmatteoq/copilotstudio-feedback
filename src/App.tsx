import { useCallback, useEffect, useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import {
  MsalProvider,
  useIsAuthenticated,
  useMsal,
} from "@azure/msal-react";
import {
  clearConfig,
  createMsalConfig,
  getDataverseScopes,
  getStoredConfig,
  saveConfig,
} from "./authConfig";
import { fetchAllFeedback, fetchEnvironmentName } from "./services/dataverse";
import { AppConfig, FeedbackItem } from "./types";
import ConfigPage from "./components/ConfigPage";
import FeedbackTable from "./components/FeedbackTable";
import "./App.css";

// ---------------------------------------------------------------------------
// Login screen (shown when MSAL is ready but user is not signed in)
// ---------------------------------------------------------------------------
function LoginPage({ config, onLogin }: { config: AppConfig; onLogin: () => void }) {
  return (
    <div className="login-container">
      <div className="login-card">
        <div className="login-logo">📊</div>
        <h2>Copilot Studio Feedback Viewer</h2>
        <p>
          Sign in with your Microsoft account to view feedback from your Copilot
          Studio agents in{" "}
          {config.environmentName ? (
            <strong>{config.environmentName}</strong>
          ) : (
            <strong>{config.dataverseUrl.replace("https://", "")}</strong>
          )}.
        </p>
        <button onClick={onLogin} className="btn-primary">
          Sign in with Microsoft
        </button>
        <button
          onClick={() => {
            clearConfig();
            window.location.reload();
          }}
          className="btn-link"
        >
          Use a different environment
        </button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Main dashboard (shown when authenticated)
// ---------------------------------------------------------------------------
function Dashboard({
  config,
  onSignOut,
}: {
  config: AppConfig;
  onSignOut: () => void;
}) {
  const { instance } = useMsal();
  const [feedbackItems, setFeedbackItems] = useState<FeedbackItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [filterText, setFilterText] = useState("");
  const [filterAgent, setFilterAgent] = useState("");
  const [environmentName, setEnvironmentName] = useState<string | undefined>(config.environmentName);

  useEffect(() => {
    void fetchEnvironmentName(instance, config).then((name) => {
      if (name) setEnvironmentName(name);
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const loadFeedback = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const items = await fetchAllFeedback(instance, config);
      setFeedbackItems(items);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load feedback.");
    } finally {
      setLoading(false);
    }
  }, [instance, config]);

  useEffect(() => {
    void loadFeedback();
  }, [loadFeedback]);

  // Unique sorted agent names for the dropdown
  const agentNames = [...new Set(feedbackItems.map((i) => i.agentName))].sort();

  // Apply filters
  const filtered = feedbackItems.filter((item) => {
    const matchAgent = !filterAgent || item.agentName === filterAgent;
    const q = filterText.toLowerCase();
    const matchText =
      !q ||
      item.feedbackText.toLowerCase().includes(q) ||
      item.agentMessage.toLowerCase().includes(q);
    return matchAgent && matchText;
  });

  const likeCount = filtered.filter((i) => i.reaction === "like").length;
  const dislikeCount = filtered.filter((i) => i.reaction === "dislike").length;

  return (
    <div className="dashboard">
      {/* ── Header ── */}
      <header className="app-header">
        <div className="header-inner">
          <span className="header-title">📊 Copilot Studio Feedback Viewer</span>
          <div className="header-actions">
            <div className="env-badge">
                {environmentName && (
                  <span className="env-name">{environmentName}</span>
                )}
                <span className="env-url">{config.dataverseUrl.replace("https://", "")}</span>
              </div>
            <button onClick={onSignOut} className="btn-secondary btn-sm">
              Sign out
            </button>
          </div>
        </div>
      </header>

      {/* ── Content ── */}
      <main className="main-content">
        {/* Filter bar */}
        <div className="filter-bar">
          <div className="filter-group">
            <label htmlFor="agent-filter">Agent</label>
            <select
              id="agent-filter"
              value={filterAgent}
              onChange={(e) => setFilterAgent(e.target.value)}
            >
              <option value="">All agents</option>
              {agentNames.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
            </select>
          </div>

          <div className="filter-group filter-group-grow">
            <label htmlFor="text-filter">Search</label>
            <input
              id="text-filter"
              type="search"
              placeholder="Search feedback text or agent message…"
              value={filterText}
              onChange={(e) => setFilterText(e.target.value)}
            />
          </div>

          <button
            className="btn-secondary btn-sm refresh-btn"
            onClick={() => void loadFeedback()}
            disabled={loading}
          >
            {loading ? "Loading…" : "↻ Refresh"}
          </button>
        </div>

        {/* Error */}
        {error && (
          <div className="error-banner" role="alert">
            <strong>Error:</strong> {error}
          </div>
        )}

        {/* Loading */}
        {loading ? (
          <div className="loading-state">
            <div className="spinner" aria-label="Loading" />
            <p>Fetching conversation transcripts…</p>
          </div>
        ) : (
          <>
            {/* Stats row */}
            {feedbackItems.length > 0 && (
              <div className="stats-row">
                <div className="stat-chip">
                  <span className="stat-label">Showing</span>
                  <span className="stat-value">
                    {filtered.length !== feedbackItems.length
                      ? `${filtered.length} / ${feedbackItems.length}`
                      : feedbackItems.length}
                  </span>
                </div>
                <div className="stat-chip stat-like">
                  <span>👍</span>
                  <span className="stat-value">{likeCount}</span>
                </div>
                <div className="stat-chip stat-dislike">
                  <span>👎</span>
                  <span className="stat-value">{dislikeCount}</span>
                </div>
              </div>
            )}

            <FeedbackTable items={filtered} />
          </>
        )}
      </main>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Authenticated wrapper (decides login vs dashboard)
// ---------------------------------------------------------------------------
function AuthenticatedApp({
  config,
  onSignOut,
}: {
  config: AppConfig;
  onSignOut: () => void;
}) {
  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const handleLogin = () => {
    void instance.loginPopup({ scopes: getDataverseScopes(config) });
  };

  if (!isAuthenticated) {
    return <LoginPage config={config} onLogin={handleLogin} />;
  }

  return <Dashboard config={config} onSignOut={onSignOut} />;
}

// ---------------------------------------------------------------------------
// Root: config gate → MSAL init → app
// ---------------------------------------------------------------------------
export default function App() {
  const [config, setConfig] = useState<AppConfig | null>(() => getStoredConfig());
  const [msalInstance, setMsalInstance] = useState<PublicClientApplication | null>(
    null
  );

  useEffect(() => {
    if (!config) {
      setMsalInstance(null);
      return;
    }
    const instance = new PublicClientApplication(createMsalConfig(config));
    instance
      .initialize()
      .then(() => setMsalInstance(instance))
      .catch((err: unknown) => console.error("MSAL init failed:", err));
  }, [config]);

  const handleSaveConfig = (newConfig: AppConfig) => {
    saveConfig(newConfig);
    setConfig(newConfig);
  };

  const handleSignOut = () => {
    clearConfig();
    setConfig(null);
    setMsalInstance(null);
  };

  if (!config) {
    return <ConfigPage onSave={handleSaveConfig} />;
  }

  if (!msalInstance) {
    return (
      <div className="loading-state loading-full">
        <div className="spinner" />
        <p>Initializing…</p>
      </div>
    );
  }

  return (
    <MsalProvider instance={msalInstance}>
      <AuthenticatedApp config={config} onSignOut={handleSignOut} />
    </MsalProvider>
  );
}
