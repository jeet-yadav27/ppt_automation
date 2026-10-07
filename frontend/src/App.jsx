import { useEffect, useState } from "react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

const STORAGE = "ftir-report";
const NAVY = "#0B3A5B";
const TEAL = "#1F7A8C";
const AMBER = "#E08A2A";
const GREEN = "#1E8C5A";
const RED = "#C0392B";

function loadSession() {
  try {
    return JSON.parse(sessionStorage.getItem(STORAGE) || "null");
  } catch {
    return null;
  }
}

async function readJson(response) {
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(body.detail || "Request failed");
  }
  return body;
}

function pct(value) {
  return `${Math.round((value || 0) * 100)}%`;
}

function ChartCard({ title, wide, children }) {
  return (
    <section className={wide ? "panel wide" : "panel"}>
      <h2>{title}</h2>
      <div style={{ width: "100%", height: 280 }}>{children}</div>
    </section>
  );
}

export default function App() {
  const existing = loadSession();
  const [name, setName] = useState(existing?.name || "");
  const [mode, setMode] = useState("server");
  const [file, setFile] = useState(null);
  const [serverSource, setServerSource] = useState("");
  const [session, setSession] = useState(existing);
  const [report, setReport] = useState(null);
  const [activity, setActivity] = useState([]);
  const [runningCount, setRunningCount] = useState(0);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    fetch("/api/health")
      .then((response) => response.json())
      .then((body) => {
        if (body.source) setServerSource(`${body.source} (${Number(body.rows).toLocaleString()} rows)`);
      })
      .catch(() => {});
  }, []);

  useEffect(() => {
    let stop = false;
    async function pollActivity() {
      try {
        const body = await readJson(await fetch("/api/activity"));
        if (!stop) {
          setActivity(body.jobs || []);
          setRunningCount(body.running || 0);
        }
      } catch {
        /* The server may still be starting. */
      }
    }
    pollActivity();
    const timer = setInterval(pollActivity, 2000);
    return () => {
      stop = true;
      clearInterval(timer);
    };
  }, []);

  useEffect(() => {
    if (!session?.id) return undefined;
    let stop = false;
    async function pollReport() {
      try {
        const body = await readJson(await fetch(`/api/reports/${session.id}`, {
          headers: { "X-Report-Token": session.token },
        }));
        if (stop) return;
        setReport(body);
        setError(body.error || "");
        if (body.status === "ready" || body.status === "error") setBusy(false);
      } catch (err) {
        if (!stop) setError(err.message);
      }
    }
    pollReport();
    const timer = setInterval(pollReport, 1000);
    return () => {
      stop = true;
      clearInterval(timer);
    };
  }, [session]);

  async function buildReport(event) {
    event.preventDefault();
    if (mode === "upload" && !file) {
      setError("Choose a CSV file, or switch back to the server dataset.");
      return;
    }
    if (file && file.size > 100 * 1024 * 1024) {
      setError("CSV must be 100 MB or smaller.");
      return;
    }
    setBusy(true);
    setError("");
    setReport(null);
    try {
      const form = new FormData();
      form.append("name", name.trim() || "Analyst");
      if (mode === "upload" && file) form.append("file", file);
      const body = await readJson(await fetch("/api/reports", {
        method: "POST",
        body: form,
      }));
      const next = { id: body.id, token: body.token, name: body.name };
      sessionStorage.setItem(STORAGE, JSON.stringify(next));
      setSession(next);
      setName(body.name);
    } catch (err) {
      setBusy(false);
      setError(err.message);
    }
  }

  async function download() {
    const response = await fetch(`/api/reports/${session.id}/download`, {
      headers: { "X-Report-Token": session.token },
    });
    if (!response.ok) {
      setError("The deck is not ready to download yet.");
      return;
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `FTIR_KPI_${(session.name || "Analyst").replace(/\s+/g, "_")}.pptx`;
    link.click();
    URL.revokeObjectURL(url);
  }

  const kpis = report?.kpis;
  const charts = report?.charts;
  const status = report?.status || (busy ? "queued" : "idle");

  return (
    <div className="app">
      <header className="topbar">
        <p className="eyebrow">Field quality · 2026</p>
        <h1>FTIR KPI desk</h1>
        <p className="lede">
          Each person gets a separate report file. Two laptops can build and download at the same time without overwriting each other.
        </p>
      </header>
      <main className="layout">
        <div className="stack">
          <form className="card composer" onSubmit={buildReport}>
            <h2>Build a deck for your laptop</h2>
            <p className="hint">Your name is printed on the title slide. An uploaded CSV is used only for your report.</p>
            <div className="choices">
              <label className="choice">
                <input type="radio" name="dataset" checked={mode === "server"} onChange={() => setMode("server")} />
                Server dataset{serverSource ? ` · ${serverSource}` : ""}
              </label>
              <label className="choice">
                <input type="radio" name="dataset" checked={mode === "upload"} onChange={() => setMode("upload")} />
                Upload my CSV
              </label>
            </div>
            {mode === "upload" && (
              <label className="file">
                FTIR extract (.csv, up to 100 MB)
                <input
                  type="file"
                  accept=".csv,text/csv"
                  onChange={(event) => setFile(event.target.files?.[0] || null)}
                />
              </label>
            )}
            <div className="row">
              <label>
                Your name
                <input value={name} onChange={(event) => setName(event.target.value)} placeholder="Analyst" maxLength={60} />
              </label>
              <button type="submit" disabled={busy}>{busy ? "Building…" : "Build my report"}</button>
              <button type="button" className="secondary" disabled={status !== "ready"} onClick={download}>Download PPT</button>
            </div>
            <div className="status">
              <span className={`dot ${status}`} />
              {status === "idle" && "Waiting for a name"}
              {status === "queued" && "Queued on the server"}
              {status === "running" && "Reading the FTIR file and writing your deck"}
              {status === "ready" && `Ready for ${session?.name}${report?.source ? ` · ${report.source}` : ""}`}
              {status === "error" && "The report failed"}
            </div>
            {error && <p className="error">{error}</p>}
          </form>

          {kpis && (
            <section className="kpis">
              <article className="card kpi"><span>FTIR rows</span><strong>{kpis.n.toLocaleString()}</strong><em>{kpis.date_min} – {kpis.date_max}</em></article>
              <article className="card kpi"><span>NG failures</span><strong>{pct(kpis.fail_rate)}</strong><em>{kpis.top_failure_plant} holds {pct(kpis.top_failure_share)}</em></article>
              <article className="card kpi"><span>Resolved</span><strong>{pct(kpis.resolution_rate)}</strong><em>{pct(kpis.resolution_0_12)} in the first 12 months</em></article>
              <article className="card kpi"><span>Judgement lag</span><strong>{Math.round(kpis.avg_lag)}d</strong><em>Median {Math.round(kpis.median_lag)} days</em></article>
            </section>
          )}

          {charts && (
            <section className="charts">
              <ChartCard title="Monthly FC OK" wide>
                <ResponsiveContainer>
                  <LineChart data={charts.monthly}>
                    <CartesianGrid stroke="#e5e8eb" vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 11 }} interval={5} />
                    <YAxis tick={{ fontSize: 11 }} />
                    <Tooltip />
                    <Line dataKey="ok" name="FC OK" stroke={GREEN} strokeWidth={2} dot={false} />
                  </LineChart>
                </ResponsiveContainer>
              </ChartCard>
              <ChartCard title="Plants">
                <ResponsiveContainer>
                  <BarChart data={charts.plants} layout="vertical" margin={{ left: 24 }}>
                    <CartesianGrid stroke="#e5e8eb" horizontal={false} />
                    <XAxis type="number" tick={{ fontSize: 11 }} />
                    <YAxis type="category" dataKey="name" width={120} tick={{ fontSize: 11 }} />
                    <Tooltip />
                    <Bar dataKey="complaints" fill={TEAL} radius={[0, 6, 6, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
              <ChartCard title="Usage buckets">
                <ResponsiveContainer>
                  <BarChart data={charts.usage}>
                    <CartesianGrid stroke="#e5e8eb" vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 10 }} interval={0} angle={-35} height={70} textAnchor="end" />
                    <YAxis tick={{ fontSize: 11 }} />
                    <Tooltip />
                    <Bar dataKey="complaints" fill={NAVY} radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
              <ChartCard title="Severity">
                <ResponsiveContainer>
                  <BarChart data={charts.ranks}>
                    <CartesianGrid stroke="#e5e8eb" vertical={false} />
                    <XAxis dataKey="rank" />
                    <YAxis />
                    <Tooltip />
                    <Bar dataKey="complaints" fill={AMBER} radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
              <ChartCard title="Dealers">
                <ResponsiveContainer>
                  <BarChart data={charts.dealers}>
                    <CartesianGrid stroke="#e5e8eb" vertical={false} />
                    <XAxis dataKey="dealer" tick={{ fontSize: 10 }} interval={0} angle={-30} height={80} textAnchor="end" />
                    <YAxis />
                    <Tooltip />
                    <Legend />
                    <Bar dataKey="sales" name="Sales" fill={NAVY} />
                    <Bar dataKey="service" name="Service" fill={AMBER} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
              <ChartCard title="Countermeasures" wide>
                <ResponsiveContainer>
                  <BarChart data={charts.measures} layout="vertical" margin={{ left: 40 }}>
                    <CartesianGrid stroke="#e5e8eb" horizontal={false} />
                    <XAxis type="number" />
                    <YAxis type="category" dataKey="name" width={180} tick={{ fontSize: 11 }} />
                    <Tooltip />
                    <Legend />
                    <Bar dataKey="yes" stackId="s" fill={GREEN} name="Yes" />
                    <Bar dataKey="partially" stackId="s" fill={AMBER} name="Partially" />
                    <Bar dataKey="no" stackId="s" fill={RED} name="No" />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
              <ChartCard title="Time to action">
                <ResponsiveContainer>
                  <BarChart data={charts.lag}>
                    <CartesianGrid stroke="#e5e8eb" vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 11 }} />
                    <YAxis />
                    <Tooltip />
                    <Bar dataKey="count" fill={NAVY} radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            </section>
          )}

          {report?.insights && (
            <section className="notes">
              <article className="card">
                <h2>Insights</h2>
                <ul>{report.insights.map((line) => <li key={line}>{line}</li>)}</ul>
              </article>
              <article className="card">
                <h2>Recommendations</h2>
                <ul>{report.recommendations.map((line) => <li key={line}>{line}</li>)}</ul>
              </article>
            </section>
          )}
        </div>
        <aside className="rail">
          <h2>On this server</h2>
          <p className="hint">{runningCount} report{runningCount === 1 ? "" : "s"} building right now. Names are visible; download links are not shared.</p>
          <ol>
            {activity.length === 0 && <li><div><div className="who">No reports yet</div><div className="meta">Be the first to build one</div></div></li>}
            {activity.map((job) => (
              <li key={job.id}>
                <div>
                  <div className="who">{job.name}</div>
                  <div className="meta">{job.created.replace("T", " ").replace("+00:00", " UTC")}</div>
                </div>
                <span className="pill">{job.status}</span>
              </li>
            ))}
          </ol>
        </aside>
      </main>
    </div>
  );
}
