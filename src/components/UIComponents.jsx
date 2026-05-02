import { useState } from "react";

// TopBar.jsx
export function TopBar({ lastRefresh, onRefresh, onConfig, totalCost, totalResources, subCount, activeView }) {
  return (
    <header className="topbar">
      <div className="topbar-left">
        <h1 className="page-title">{activeView}</h1>
        {lastRefresh && (
          <span className="refresh-time">
            Updated {lastRefresh.toLocaleTimeString()}
          </span>
        )}
      </div>
      <div className="topbar-right">
        <div className="topbar-stat">
          <span className="ts-label">MTD Spend</span>
          <span className="ts-value green">
            ${totalCost?.toLocaleString("en-US", { maximumFractionDigits: 0 })}
          </span>
        </div>
        <div className="topbar-stat">
          <span className="ts-label">Resources</span>
          <span className="ts-value blue">{totalResources?.toLocaleString()}</span>
        </div>
        <button className="btn-icon" onClick={onRefresh} title="Refresh data">
          ↻
        </button>
        <button className="btn-icon" onClick={onConfig} title="Settings">
          ⚙
        </button>
      </div>
    </header>
  );
}

// Sidebar.jsx
export function Sidebar({ views, activeView, onViewChange, subscriptions }) {
  return (
    <aside className="sidebar">
      <div className="sidebar-logo">
        <div className="logo-mark">A</div>
        <div className="logo-text">
          <span className="logo-name">AzureReader</span>
          <span className="logo-tenant">avontus.com</span>
        </div>
      </div>
      <nav className="sidebar-nav">
        {views.map((v) => (
          <button
            key={v}
            className={`nav-item ${activeView === v ? "active" : ""}`}
            onClick={() => onViewChange(v)}
          >
            <span className="nav-icon">
              {v === "Overview" && "◉"}
              {v === "Costs" && "💳"}
              {v === "Resources" && "🗄"}
              {v === "Subscriptions" && "📋"}
            </span>
            {v}
          </button>
        ))}
      </nav>
      <div className="sidebar-section-label">SUBSCRIPTIONS</div>
      <div className="sidebar-subs">
        {subscriptions?.slice(0, 8).map((s) => (
          <div key={s.subscriptionId} className="sub-pill">
            <span className="sub-dot" />
            <span className="sub-name">{s.displayName}</span>
          </div>
        ))}
        {subscriptions?.length > 8 && (
          <div className="sub-more">+{subscriptions.length - 8} more</div>
        )}
      </div>
      <div className="sidebar-footer">
        <div className="sf-tenant">Tenant: bd98204b</div>
        <div className="sf-app">App: AzureReader</div>
      </div>
    </aside>
  );
}

// LoadingScreen.jsx
export function LoadingScreen() {
  return (
    <div className="fullscreen-center">
      <div className="loader-wrap">
        <div className="azure-logo-anim">
          <div className="ring r1" />
          <div className="ring r2" />
          <div className="ring r3" />
          <div className="ring-center">A</div>
        </div>
        <div className="loader-text">Fetching Azure data…</div>
        <div className="loader-steps">
          <Step label="Authenticating with Azure AD" />
          <Step label="Loading subscriptions" delay={0.4} />
          <Step label="Querying Cost Management" delay={0.8} />
          <Step label="Running Resource Graph queries" delay={1.2} />
        </div>
      </div>
    </div>
  );
}

function Step({ label, delay = 0 }) {
  return (
    <div className="loader-step" style={{ animationDelay: `${delay}s` }}>
      <span className="step-dot" />
      {label}
    </div>
  );
}

// ErrorScreen.jsx
export function ErrorScreen({ error, onRetry, onConfig }) {
  return (
    <div className="fullscreen-center">
      <div className="error-card">
        <div className="error-icon">⚠</div>
        <h2 className="error-title">Connection Error</h2>
        <p className="error-msg">{error}</p>
        <div className="error-hints">
          <p>Common causes:</p>
          <ul>
            <li>Client secret expired or incorrect</li>
            <li>App registration missing API permissions</li>
            <li>CORS restriction (use Azure Functions proxy in production)</li>
            <li>Subscription reader role not assigned</li>
          </ul>
        </div>
        <div className="error-actions">
          <button className="btn-primary" onClick={onRetry}>
            Retry
          </button>
          <button className="btn-secondary" onClick={onConfig}>
            Update Secret
          </button>
        </div>
      </div>
    </div>
  );
}

// ConfigModal.jsx
export function ConfigModal({ onSave, onClose, hasSecret }) {
  const [val, setVal] = useState("");
  return (
    <div className="modal-overlay">
      <div className="modal-card">
        <div className="modal-header">
          <h2>Azure Configuration</h2>
          {hasSecret && (
            <button className="btn-icon" onClick={onClose}>
              ✕
            </button>
          )}
        </div>
        <div className="modal-body">
          <div className="config-info">
            <div className="cfg-row">
              <span className="cfg-label">Tenant ID</span>
              <code>bd98204b-b981-4d03-8796-356d537927eb</code>
            </div>
            <div className="cfg-row">
              <span className="cfg-label">Client ID</span>
              <code>3977e66a-cdf1-419d-9d0d-70e8cf3a76ed</code>
            </div>
            <div className="cfg-row">
              <span className="cfg-label">Cert Expiry</span>
              <code>3/5/2028</code>
            </div>
          </div>
          <label className="field-label">
            Client Secret
            <input
              type="password"
              className="field-input"
              placeholder="Paste client secret value…"
              value={val}
              onChange={(e) => setVal(e.target.value)}
            />
          </label>
          <p className="modal-note">
            ⚠ In production, set <code>VITE_AZURE_CLIENT_SECRET</code> as an environment variable
            in Azure Static Web Apps. Never commit secrets to source control.
          </p>
        </div>
        <div className="modal-footer">
          <button
            className="btn-primary"
            disabled={!val}
            onClick={() => onSave(val)}
          >
            Connect
          </button>
        </div>
      </div>
    </div>
  );
}

export default { TopBar, Sidebar, LoadingScreen, ErrorScreen, ConfigModal };
