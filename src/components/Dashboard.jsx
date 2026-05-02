import { useState } from "react";
import { SubscriptionCards } from "./DataComponents";
import CostChart from "./CostChart";
import { ResourceTable } from "./DataComponents";
import { ResourceByType } from "./DataComponents";
import { CostBreakdown } from "./DataComponents";
import { TopBar } from "./UIComponents";
import { Sidebar } from "./UIComponents";

const VIEWS = ["Overview", "Costs", "Resources", "Subscriptions"];

export default function Dashboard({ data, lastRefresh, onRefresh, onConfig }) {
  const [activeView, setActiveView] = useState("Overview");

  const totalCost = data.costs.reduce((sum, c) => {
    const rows = c.data?.properties?.rows || [];
    return sum + rows.reduce((s, r) => s + (r[0] || 0), 0);
  }, 0);

  const totalResources = data.resourcesByType?.data?.rows?.reduce(
    (s, r) => s + (r[1] || 0),
    0
  ) || 0;

  return (
    <div className="dashboard">
      <Sidebar
        views={VIEWS}
        activeView={activeView}
        onViewChange={setActiveView}
        subscriptions={data.subscriptions}
      />
      <div className="main-content">
        <TopBar
          lastRefresh={lastRefresh}
          onRefresh={onRefresh}
          onConfig={onConfig}
          totalCost={totalCost}
          totalResources={totalResources}
          subCount={data.subscriptions.length}
          activeView={activeView}
        />
        <div className="content-area">
          {activeView === "Overview" && (
            <OverviewView data={data} totalCost={totalCost} totalResources={totalResources} />
          )}
          {activeView === "Costs" && <CostsView data={data} />}
          {activeView === "Resources" && <ResourcesView data={data} />}
          {activeView === "Subscriptions" && <SubscriptionsView data={data} />}
        </div>
      </div>
    </div>
  );
}

function OverviewView({ data, totalCost, totalResources }) {
  return (
    <div className="view-grid">
      <div className="grid-row summary-row">
        <StatCard
          label="Total MTD Spend"
          value={`$${totalCost.toLocaleString("en-US", { maximumFractionDigits: 0 })}`}
          sub={`Across ${data.subscriptions.length} subscriptions`}
          accent="#4ade80"
          icon="💰"
        />
        <StatCard
          label="Total Resources"
          value={totalResources.toLocaleString()}
          sub={`${data.resourcesByType?.data?.rows?.length || 0} resource types`}
          accent="#60a5fa"
          icon="🗄️"
        />
        <StatCard
          label="Subscriptions"
          value={data.subscriptions.length}
          sub="Active in tenant"
          accent="#f59e0b"
          icon="📋"
        />
        <StatCard
          label="Resource Groups"
          value={data.resourcesByRG?.data?.rows?.length || 0}
          sub="Total groups"
          accent="#a78bfa"
          icon="📁"
        />
      </div>
      <div className="grid-row two-col">
        <div className="card">
          <h3 className="card-title">Monthly Cost by Subscription</h3>
          <CostBreakdown costs={data.costs} subscriptions={data.subscriptions} />
        </div>
        <div className="card">
          <h3 className="card-title">Resources by Type</h3>
          <ResourceByType data={data.resourcesByType} />
        </div>
      </div>
      <div className="grid-row">
        <div className="card full-width">
          <h3 className="card-title">12-Month Cost Trend</h3>
          <CostChart trends={data.trends} subscriptions={data.subscriptions} />
        </div>
      </div>
      <div className="grid-row">
        <div className="card full-width">
          <h3 className="card-title">Resources by Location</h3>
          <LocationGrid data={data.resourcesByLocation} />
        </div>
      </div>
    </div>
  );
}

function CostsView({ data }) {
  return (
    <div className="view-grid">
      <div className="grid-row">
        <div className="card full-width">
          <h3 className="card-title">12-Month Cost Trend (All Subscriptions)</h3>
          <CostChart trends={data.trends} subscriptions={data.subscriptions} large />
        </div>
      </div>
      <div className="grid-row">
        <div className="card full-width">
          <h3 className="card-title">Cost Breakdown by Service & Resource Group</h3>
          <CostBreakdownTable costs={data.costs} subscriptions={data.subscriptions} />
        </div>
      </div>
    </div>
  );
}

function ResourcesView({ data }) {
  return (
    <div className="view-grid">
      <div className="grid-row two-col">
        <div className="card">
          <h3 className="card-title">Resources by Type</h3>
          <ResourceByType data={data.resourcesByType} full />
        </div>
        <div className="card">
          <h3 className="card-title">Resources by Location</h3>
          <LocationGrid data={data.resourcesByLocation} />
        </div>
      </div>
      <div className="grid-row">
        <div className="card full-width">
          <h3 className="card-title">Resource Groups</h3>
          <ResourceTable data={data.resourcesByRG} />
        </div>
      </div>
    </div>
  );
}

function SubscriptionsView({ data }) {
  return (
    <div className="view-grid">
      <div className="grid-row">
        <div className="card full-width">
          <h3 className="card-title">All Subscriptions</h3>
          <SubscriptionCards
            subscriptions={data.subscriptions}
            costs={data.costs}
            resourcesByRG={data.resourcesByRG}
          />
        </div>
      </div>
    </div>
  );
}

function StatCard({ label, value, sub, accent, icon }) {
  return (
    <div className="stat-card" style={{ "--accent": accent }}>
      <div className="stat-icon">{icon}</div>
      <div className="stat-body">
        <div className="stat-value">{value}</div>
        <div className="stat-label">{label}</div>
        <div className="stat-sub">{sub}</div>
      </div>
      <div className="stat-bar" />
    </div>
  );
}

function LocationGrid({ data }) {
  const rows = data?.data?.rows || [];
  const max = rows.reduce((m, r) => Math.max(m, r[1] || 0), 0);
  return (
    <div className="location-grid">
      {rows.slice(0, 12).map((r, i) => (
        <div key={i} className="location-item">
          <div className="loc-name">{r[0] || "Unknown"}</div>
          <div className="loc-bar-wrap">
            <div
              className="loc-bar"
              style={{ width: `${max ? ((r[1] / max) * 100).toFixed(1) : 0}%` }}
            />
          </div>
          <div className="loc-count">{r[1]?.toLocaleString()}</div>
        </div>
      ))}
      {rows.length === 0 && <div className="empty-msg">No location data available</div>}
    </div>
  );
}

function CostBreakdownTable({ costs, subscriptions }) {
  const rows = [];
  for (const c of costs) {
    const subName = c.subscription?.displayName || c.subscription?.subscriptionId;
    for (const r of c.data?.properties?.rows || []) {
      rows.push({
        cost: r[0] || 0,
        service: r[1] || "Unknown",
        rg: r[2] || "Unknown",
        sub: subName,
        currency: r[3] || "USD",
      });
    }
  }
  rows.sort((a, b) => b.cost - a.cost);

  return (
    <div className="data-table-wrap">
      <table className="data-table">
        <thead>
          <tr>
            <th>Subscription</th>
            <th>Service</th>
            <th>Resource Group</th>
            <th className="num">Cost (MTD)</th>
          </tr>
        </thead>
        <tbody>
          {rows.slice(0, 100).map((r, i) => (
            <tr key={i}>
              <td>
                <span className="badge blue">{r.sub}</span>
              </td>
              <td>{r.service}</td>
              <td>{r.rg}</td>
              <td className="num cost-cell">
                ${r.cost.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </td>
            </tr>
          ))}
          {rows.length === 0 && (
            <tr>
              <td colSpan={4} className="empty-msg">
                No cost data available for this period
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}
