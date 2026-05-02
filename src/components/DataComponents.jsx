const TYPE_COLORS = [
  "#4ade80","#60a5fa","#f59e0b","#a78bfa",
  "#f87171","#34d399","#fb923c","#38bdf8",
  "#e879f9","#fbbf24","#6ee7b7","#93c5fd",
];

// ResourceByType
export function ResourceByType({ data, full }) {
  const rows = data?.data?.rows || [];
  const limit = full ? rows.length : 10;
  const max = rows.reduce((m, r) => Math.max(m, r[1] || 0), 0);

  if (!rows.length) return <div className="empty-msg">No resource type data</div>;

  return (
    <div className="res-type-list">
      {rows.slice(0, limit).map((r, i) => {
        const pct = max ? ((r[1] / max) * 100).toFixed(1) : 0;
        const label = r[0]?.split("/").pop() || "unknown";
        return (
          <div key={i} className="res-type-row">
            <span className="res-dot" style={{ background: TYPE_COLORS[i % TYPE_COLORS.length] }} />
            <span className="res-type-name" title={r[0]}>{label}</span>
            <div className="res-bar-wrap">
              <div className="res-bar" style={{ width: `${pct}%`, background: TYPE_COLORS[i % TYPE_COLORS.length] }} />
            </div>
            <span className="res-count">{r[1]?.toLocaleString()}</span>
          </div>
        );
      })}
    </div>
  );
}

// ResourceTable (resource groups)
export function ResourceTable({ data }) {
  const rows = data?.data?.rows || [];
  if (!rows.length) return <div className="empty-msg">No resource group data</div>;

  return (
    <div className="data-table-wrap">
      <table className="data-table">
        <thead>
          <tr>
            <th>Resource Group</th>
            <th>Subscription ID</th>
            <th className="num">Resources</th>
            <th className="num">Types</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i}>
              <td>{r[0]}</td>
              <td><code className="sub-code">{r[1]?.slice(0, 8)}…</code></td>
              <td className="num">{r[2]?.toLocaleString()}</td>
              <td className="num">{r[3]?.toLocaleString()}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// CostBreakdown (donut-style list)
export function CostBreakdown({ costs, subscriptions }) {
  const subTotals = costs.map((c) => {
    const rows = c.data?.properties?.rows || [];
    const total = rows.reduce((s, r) => s + (r[0] || 0), 0);
    return { name: c.subscription?.displayName || c.subscription?.subscriptionId, total };
  }).sort((a, b) => b.total - a.total);

  const grandTotal = subTotals.reduce((s, t) => s + t.total, 0);

  if (!subTotals.length) return <div className="empty-msg">No cost data available</div>;

  return (
    <div className="cost-breakdown">
      {subTotals.map((s, i) => {
        const pct = grandTotal ? ((s.total / grandTotal) * 100).toFixed(1) : 0;
        return (
          <div key={i} className="cb-row">
            <span className="cb-dot" style={{ background: TYPE_COLORS[i % TYPE_COLORS.length] }} />
            <span className="cb-name">{s.name}</span>
            <div className="cb-bar-wrap">
              <div className="cb-bar" style={{ width: `${pct}%`, background: TYPE_COLORS[i % TYPE_COLORS.length] }} />
            </div>
            <span className="cb-pct">{pct}%</span>
            <span className="cb-val">${s.total.toLocaleString("en-US", { maximumFractionDigits: 0 })}</span>
          </div>
        );
      })}
    </div>
  );
}

// SubscriptionCards
export function SubscriptionCards({ subscriptions, costs, resourcesByRG }) {
  return (
    <div className="sub-cards-grid">
      {subscriptions.map((sub, i) => {
        const costEntry = costs.find((c) => c.subscription?.subscriptionId === sub.subscriptionId);
        const rows = costEntry?.data?.properties?.rows || [];
        const total = rows.reduce((s, r) => s + (r[0] || 0), 0);
        const services = [...new Set(rows.map((r) => r[1]).filter(Boolean))].slice(0, 5);

        const rgRows = resourcesByRG?.data?.rows?.filter((r) => r[1] === sub.subscriptionId) || [];

        return (
          <div key={sub.subscriptionId} className="sub-card">
            <div className="sc-header">
              <div className="sc-color" style={{ background: TYPE_COLORS[i % TYPE_COLORS.length] }} />
              <div>
                <div className="sc-name">{sub.displayName}</div>
                <div className="sc-id">{sub.subscriptionId}</div>
              </div>
            </div>
            <div className="sc-state">
              <span className={`badge ${sub.state === "Enabled" ? "green" : "red"}`}>
                {sub.state || "Unknown"}
              </span>
            </div>
            <div className="sc-metrics">
              <div className="sc-metric">
                <div className="scm-val">${total.toLocaleString("en-US", { maximumFractionDigits: 0 })}</div>
                <div className="scm-label">MTD Cost</div>
              </div>
              <div className="sc-metric">
                <div className="scm-val">{rgRows.length}</div>
                <div className="scm-label">Res. Groups</div>
              </div>
              <div className="sc-metric">
                <div className="scm-val">{services.length}</div>
                <div className="scm-label">Services</div>
              </div>
            </div>
            {services.length > 0 && (
              <div className="sc-services">
                {services.map((s, j) => (
                  <span key={j} className="service-tag">{s}</span>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

export default { ResourceByType, ResourceTable, CostBreakdown, SubscriptionCards };
