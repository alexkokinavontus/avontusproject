const COLORS = [
  "#4ade80", "#60a5fa", "#f59e0b", "#a78bfa",
  "#f87171", "#34d399", "#fb923c", "#38bdf8",
];

export default function CostChart({ trends, subscriptions, large }) {
  // Build monthly buckets across all subscriptions
  const monthMap = {};

  trends.forEach((t, idx) => {
    const rows = t.data?.properties?.rows || [];
    const subName = t.subscription?.displayName?.split(" ")[0] || `Sub${idx + 1}`;
    rows.forEach((r) => {
      const rawDate = r[1];
      if (!rawDate) return;
      const d = new Date(rawDate);
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      if (!monthMap[key]) monthMap[key] = { label: key, total: 0 };
      monthMap[key][subName] = (monthMap[key][subName] || 0) + (r[0] || 0);
      monthMap[key].total += r[0] || 0;
    });
  });

  const months = Object.values(monthMap).sort((a, b) => a.label.localeCompare(b.label)).slice(-12);
  const subNames = [...new Set(trends.map((t, i) => t.subscription?.displayName?.split(" ")[0] || `Sub${i + 1}`))];
  const maxVal = months.reduce((m, mo) => Math.max(m, mo.total), 0);

  const W = large ? 900 : 600;
  const H = large ? 280 : 200;
  const PADL = 60, PADR = 20, PADT = 20, PADB = 50;
  const chartW = W - PADL - PADR;
  const chartH = H - PADT - PADB;
  const barW = months.length ? (chartW / months.length) * 0.7 : 30;
  const gap = months.length ? chartW / months.length : 40;

  if (months.length === 0) {
    return <div className="empty-msg">No trend data available</div>;
  }

  return (
    <div className="chart-wrap">
      <svg viewBox={`0 0 ${W} ${H}`} className="cost-chart">
        {/* Grid lines */}
        {[0, 0.25, 0.5, 0.75, 1].map((pct) => {
          const y = PADT + chartH * (1 - pct);
          const val = maxVal * pct;
          return (
            <g key={pct}>
              <line x1={PADL} y1={y} x2={W - PADR} y2={y} stroke="#ffffff15" strokeWidth={1} />
              <text x={PADL - 6} y={y + 4} textAnchor="end" fontSize={10} fill="#6b7280">
                {val >= 1000 ? `$${(val / 1000).toFixed(0)}k` : `$${val.toFixed(0)}`}
              </text>
            </g>
          );
        })}

        {/* Stacked bars */}
        {months.map((mo, i) => {
          const x = PADL + i * gap + gap / 2 - barW / 2;
          let yOffset = PADT + chartH;
          return (
            <g key={mo.label}>
              {subNames.map((sn, si) => {
                const val = mo[sn] || 0;
                if (!val) return null;
                const h = maxVal ? (val / maxVal) * chartH : 0;
                yOffset -= h;
                return (
                  <rect
                    key={sn}
                    x={x}
                    y={yOffset}
                    width={barW}
                    height={h}
                    fill={COLORS[si % COLORS.length]}
                    opacity={0.85}
                    rx={2}
                  >
                    <title>{`${sn}: $${val.toLocaleString("en-US", { maximumFractionDigits: 0 })}`}</title>
                  </rect>
                );
              })}
              <text
                x={x + barW / 2}
                y={PADT + chartH + 16}
                textAnchor="middle"
                fontSize={9}
                fill="#9ca3af"
              >
                {mo.label.slice(5)}/{mo.label.slice(2, 4)}
              </text>
            </g>
          );
        })}
      </svg>

      {/* Legend */}
      <div className="chart-legend">
        {subNames.slice(0, 6).map((sn, i) => (
          <div key={sn} className="legend-item">
            <span className="legend-dot" style={{ background: COLORS[i % COLORS.length] }} />
            {sn}
          </div>
        ))}
      </div>
    </div>
  );
}
