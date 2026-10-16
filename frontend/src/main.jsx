import React, { useState, useEffect, useRef } from "react";
import { createRoot } from "react-dom/client";
import { Activity, ArrowRight, BarChart3, CheckCircle2, Clock3, FileSpreadsheet, History, LayoutDashboard, LogOut, Menu, Settings, ShieldCheck, UploadCloud, UserCircle2, X } from "lucide-react";
import "./styles.css";
import "./auth.css";
import "./workspace.css";

const API = import.meta.env.VITE_API_URL || "";
const services = [
  { id: "arrears-difference", title: "Loan Arrears Difference", icon: Activity, purpose: "Compare beginning and current arrears, match officers to targets, and rank changes in a formatted workbook.", required: "Two dated arrears CSVs + Targets.xlsx", output: "Arrears_Report.xlsx", endpoint: "/api/loan-arrears-difference", fields: ["first_file", "second_file", "targets_file"] },
  { id: "all-loans", title: "Clean All Loans", icon: FileSpreadsheet, purpose: "Standardize loan records, dates, branches, money fields, and remove summary rows.", required: "One CSV loan report", output: "AllLoans_Clean.csv", endpoint: "/api/clean-all-loans", fields: ["file"] },
  { id: "portfolio-risk", title: "Portfolio Risk", icon: ShieldCheck, purpose: "Prepare a consistent portfolio-at-risk report for analysis and review.", required: "One CSV portfolio risk report", output: "PortfolioRiskReport_Clean.csv", endpoint: "/api/clean-portfolio-risk", fields: ["file"] },
  { id: "loan-arrears", title: "Loan Arrears", icon: Activity, purpose: "Clean arrears records and normalize balances, officers, branches, and dates.", required: "One CSV loan arrears report", output: "LoanArrears_Clean.csv", endpoint: "/api/clean-loan-arrears", fields: ["file"] },
  { id: "ro-dashboard", title: "RO Performance Dashboard", icon: BarChart3, purpose: "Match officers to targets and generate a weekly and month-to-date performance workbook.", required: "Loan CSV + Targets.xlsx", output: "RO_Performance_Dashboard.xlsx", endpoint: "/api/ro-performance-dashboard", fields: ["loan_file", "targets_file"] },
  { id: "par-performance", title: "PAR Performance", icon: BarChart3, purpose: "Rank active loan officers by portfolio-at-risk performance and produce a company total report.", required: "Portfolio Risk CSV + Targets.xlsx", output: "PAR_Performance_Report.xlsx", endpoint: "/api/par-performance", fields: ["par_file", "targets_file"] },
  { id: "repayment-report", title: "Power BI Repayments", icon: FileSpreadsheet, purpose: "Clean and enrich a loan repayment report for direct Power BI use, with monthly revenue fields.", required: "Repayment CSV or Excel report", output: "Cleaned_LoanRepaymentReport.csv", endpoint: "/api/clean-repayment-report", fields: ["file"] },
  { id: "ro-customer-mix", title: "RO Customer Mix", icon: BarChart3, purpose: "Measure monthly RO disbursement performance alongside new and repeat customer contribution.", required: "Loan CSV + Targets.xlsx", output: "RO_Monthly_Performance_Customer_Mix.xlsx", endpoint: "/api/ro-monthly-customer-mix", fields: ["loan_file", "targets_file"] },
];

function App() {
  const [page, setPage] = useState("workspace");
  const [active, setActive] = useState(null);
  const [jobs, setJobs] = useState([]);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const addJob = (job) => setJobs((current) => [job, ...current]);
  useEffect(() => {
    if (!sidebarOpen) return;
    const close = (event) => { if (event.key === "Escape") setSidebarOpen(false); };
    document.addEventListener("keydown", close);
    return () => document.removeEventListener("keydown", close);
  }, [sidebarOpen]);
  return <div className="app-shell">
    <a className="skip-link" href="#main-content">Skip to workspace</a>
    {sidebarOpen && <button className="navigation-overlay" aria-label="Dismiss navigation" onClick={() => setSidebarOpen(false)} />}
    <aside className={sidebarOpen ? "sidebar open" : "sidebar"}>
      <div className="brand"><div className="brand-mark">E</div><div><strong>EasyAnalysis</strong><span>by Panta Development and AI</span></div><button className="icon-btn mobile-close" aria-label="Close navigation" onClick={() => setSidebarOpen(false)}><X size={18} /></button></div>
      <div className="workspace-label">WORKSPACE</div>
      <nav>{[["workspace", LayoutDashboard, "Workspace"], ["history", History, "History"], ["profile", UserCircle2, "Profile"], ["settings", Settings, "Settings"]].map(([id, Icon, label]) => <button className={page === id ? "nav-item active" : "nav-item"} aria-current={page === id ? "page" : undefined} key={id} onClick={() => { setPage(id); setSidebarOpen(false); }}><Icon size={18} />{label}</button>)}</nav>
      <div className="sidebar-footer"><div className="plan"><div className="plan-icon"><ShieldCheck size={17} /></div><div><b>Business workspace</b><span>Secure processing enabled</span></div></div><button className="nav-item"><LogOut size={18} />Sign out</button></div>
    </aside>
    <main className="main" id="main-content" tabIndex={-1}><header className="topbar"><button className="icon-btn menu-btn" aria-label="Open navigation" onClick={() => setSidebarOpen(true)}><Menu size={20} /></button><div className="breadcrumbs"><span>EasyAnalysis</span><ArrowRight size={14} /><b>{page[0].toUpperCase() + page.slice(1)}</b></div><button className="user user-button" aria-label="Open profile" onClick={() => setPage("profile")}><div><b>Data Team</b><span>Your workspace</span></div><div className="avatar">DT</div></button></header>
      {page === "workspace" && <Workspace onSelect={setActive} jobs={jobs} onHistory={() => setPage("history")} />}
      {page === "history" && <HistoryPage jobs={jobs} />}
      {page === "profile" && <ProfilePage />}
      {page === "settings" && <SettingsPage />}
    </main>
    {active && <UploadModal service={active} onClose={() => setActive(null)} onCompleted={(job) => { addJob(job); setActive(null); }} />}
  </div>;
}

function AuthScreen({ onAuthenticated }) { const [mode, setMode] = useState("login"); const [form, setForm] = useState({ company_name: "", full_name: "", email: "", password: "" }); const [error, setError] = useState(""); const [busy, setBusy] = useState(false); const submit = async (event) => { event.preventDefault(); setBusy(true); setError(""); try { const response = await fetch(`${API}/api/auth/${mode}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(form) }); const body = await response.json(); if (!response.ok) throw new Error(body.detail || "Unable to authenticate."); onAuthenticated(body); } catch (e) { setError(e.message); } finally { setBusy(false); } }; return <div className="auth-shell"><div className="auth-brand"><div className="brand-mark">E</div><div><strong>EasyAnalysis</strong><span>by Panta Development and AI</span></div></div><div className="auth-copy"><p className="eyebrow">OPERATIONS INTELLIGENCE</p><h1>Reports that move your business forward.</h1><p>Securely clean, standardize, and understand your lending operations from one focused workspace.</p><div className="auth-proof"><CheckCircle2 size={17} /> Built for finance and credit teams</div></div><form className="auth-card" onSubmit={submit}><p className="eyebrow">{mode === "login" ? "WELCOME BACK" : "CREATE WORKSPACE"}</p><h2>{mode === "login" ? "Sign in to EasyAnalysis" : "Create your company account"}</h2><p className="muted">{mode === "login" ? "Continue to your operations workspace." : "Start processing reports securely."}</p>{mode === "register" && <><label>Company name<input required value={form.company_name} onChange={(e) => setForm({ ...form, company_name: e.target.value })} placeholder="Your company" /></label><label>Your name<input required value={form.full_name} onChange={(e) => setForm({ ...form, full_name: e.target.value })} placeholder="Full name" /></label></>}<label>Email address<input required type="email" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} placeholder="you@company.com" /></label><label>Password<input required minLength="8" type="password" value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} placeholder="At least 8 characters" /></label>{error && <div className="error-box">{error}</div>}<button className="primary-btn auth-submit" disabled={busy}>{busy ? "Working..." : mode === "login" ? "Sign in" : "Create account"} <ArrowRight size={16} /></button><button type="button" className="switch-auth" onClick={() => { setMode(mode === "login" ? "register" : "login"); setError(""); }}>{mode === "login" ? "New to EasyAnalysis? Create a company account" : "Already have an account? Sign in"}</button></form></div> }

function Workspace({ onSelect, jobs, onHistory }) {
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState("all");
  const isCleaner = (service) => service.endpoint.includes("/clean-");
  const visible = services.filter((service) =>
    `${service.title} ${service.purpose}`.toLowerCase().includes(query.toLowerCase()) &&
    (category === "all" || (category === "clean" ? isCleaner(service) : !isCleaner(service)))
  );
  return <section className="content">
    <div className="hero-row"><div><p className="eyebrow">YOUR ANALYSIS WORKSPACE</p><h1>Less preparation.<br /><span>More perspective.</span></h1><p className="muted">From everyday loan reports to a clearer view of performance.<br />Choose your analysis. Upload your files. Get straight to the results.</p></div><div className="hero-note"><div className="hero-note-icon"><BarChart3 size={32} /></div><span>BUILT FOR YOUR WORKDAY</span><b>Clear reports.<br />Confident decisions.</b><small>Cleaning, comparisons and performance<br />in one focused workspace.</small></div></div>
    <div className="summary-grid"><Metric label="Completed this session" value={jobs.length} icon={CheckCircle2} /><Metric label="Available workflows" value={services.length} icon={BarChart3} /><Metric label="Report preparation" value="Automatic title check" icon={FileSpreadsheet} /></div>
    <details className="preparation-note"><summary><ShieldCheck size={18} /><span>Report preparation, taken care of</span><span className="preparation-hint">How it works</span></summary><p>We check for a report title above your column labels and skip it before cleaning. Files that already start with column labels stay intact. Targets workbooks keep their original structure.</p></details>
    <div className="section-heading"><div><p className="eyebrow">TOOLS FOR YOUR TEAM</p><h2>What are we working on today?</h2><p className="muted">Find the right tool for your next report.</p></div><label className="service-search"><span className="sr-only">Search workflows</span><input type="search" value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search workflows..." /></label></div>
    <div className="workflow-tabs" aria-label="Workflow categories">{[["all", "All workflows"], ["clean", "Report cleaning"], ["analysis", "Analysis & dashboards"]].map(([id, label]) => <button key={id} className={category === id ? "selected" : ""} aria-pressed={category === id} onClick={() => setCategory(id)}>{label}</button>)}<span>{visible.length} workflows</span></div>
    <div className="service-grid" aria-live="polite">{visible.map((service) => <ServiceCard key={service.id} service={service} onSelect={onSelect} />)}</div>
    {!visible.length && <div className="empty-state"><b>No matching workflows</b><small>Try another search or category.</small></div>}
    <div className="recent-heading"><h2>Recent activity</h2><button className="text-btn" onClick={onHistory}>View history <ArrowRight size={15} /></button></div>{jobs.length ? <div className="activity-table">{jobs.slice(0, 3).map((job) => <div className="activity-row" key={job.id}><div className="activity-icon"><CheckCircle2 size={18} /></div><div><b>{job.title}</b><span>{job.output}</span></div><span className="status-pill">Completed</span><small>{job.time}</small></div>)}</div> : <div className="empty-state"><History size={22} /><span>Your next insight starts here</span><small>Completed reports will appear here during this session.</small></div>}
  </section>;
}
function Metric({ label, value, icon: Icon }) { return <div className="metric"><div className="metric-icon"><Icon size={18} /></div><div><span>{label}</span><strong>{value}</strong></div></div> }
function ServiceCard({ service, onSelect }) { const Icon = service.icon; return <article className="service-card"><div className="service-top"><div className="service-icon"><Icon size={21} /></div><span className={service.id === "arrears-difference" ? "service-index new-analysis" : "service-index"}>{service.id === "arrears-difference" ? "NEW ANALYSIS" : service.endpoint.includes("/clean-") ? "CLEANING" : "ANALYSIS"}</span></div><h3>{service.title}</h3><p>{service.purpose}</p><div className="service-detail"><span>Required</span><b>{service.required}</b></div><div className="service-detail"><span>Output</span><b>{service.output}</b></div><button className="primary-btn" onClick={() => onSelect(service)}>Start processing <ArrowRight size={16} /></button></article> }

function UploadModal({ service, onClose, onCompleted }) {
  const [files, setFiles] = useState({});
  const [state, setState] = useState("idle");
  const [error, setError] = useState("");
  const dialog = useRef(null);
  const busy = state !== "idle" && state !== "completed";
  const busyRef = useRef(busy);
  busyRef.current = busy;
  useEffect(() => {
    const previousFocus = document.activeElement;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    dialog.current?.focus();
    const handleKey = (event) => {
      if (event.key === "Escape" && !busyRef.current) onClose();
      if (event.key !== "Tab") return;
      const elements = [...dialog.current.querySelectorAll('button:not(:disabled), input:not(:disabled), [tabindex="0"]')];
      const first = elements[0];
      const last = elements[elements.length - 1];
      if (!first) { event.preventDefault(); return; }
      if (event.shiftKey && (document.activeElement === first || document.activeElement === dialog.current)) { event.preventDefault(); last.focus(); }
      else if (!event.shiftKey && (document.activeElement === last || document.activeElement === dialog.current)) { event.preventDefault(); first.focus(); }
    };
    document.addEventListener("keydown", handleKey);
    return () => { document.body.style.overflow = previousOverflow; document.removeEventListener("keydown", handleKey); previousFocus?.focus(); };
  }, []);
  const pick = (key, event) => {
    const file = event.target.files[0];
    setFiles((current) => ({ ...current, [key]: file }));
    setError("");
  };
  const submit = async () => {
    if (service.fields.some((field) => !files[field])) { setError("Please select all required files before processing."); return; }
    setError(""); setState("uploading");
    const form = new FormData();
    service.fields.forEach((field) => form.append(field, files[field]));
    try {
      setState("processing");
      const response = await fetch(`${API}${service.endpoint}`, { method: "POST", body: form });
      if (!response.ok) { const body = await response.json().catch(() => ({})); throw new Error(body.detail || "The report could not be processed."); }
      setState("generating");
      const blob = await response.blob();
      const disposition = response.headers.get("content-disposition") || "";
      const name = disposition.match(/filename="?([^";]+)"?/)?.[1] || service.output;
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a"); link.href = url; link.download = name; link.click(); URL.revokeObjectURL(url);
      setState("completed");
      setTimeout(() => onCompleted({ id: crypto.randomUUID(), title: service.title, output: name, time: "Just now" }), 500);
    } catch (e) { setState("idle"); setError(e.message); }
  };
  const selected = service.fields.filter((field) => files[field]).length;
  return <div className="modal-backdrop"><div className="modal" ref={dialog} tabIndex={-1} role="dialog" aria-modal="true" aria-labelledby="upload-title">
    <div className="modal-head"><div><p className="eyebrow">PREPARE YOUR REPORT</p><h2 id="upload-title">{service.title}</h2></div><button className="icon-btn" aria-label="Close upload" disabled={busy} onClick={onClose}><X size={20} /></button></div>
    {state === "idle" ? <>
      <p className="muted">{service.purpose}</p>
      <div className="upload-overview"><span><FileSpreadsheet size={16} /> {service.output.endsWith(".csv") ? "CSV download" : "Excel workbook"}</span><span>{service.fields.length} source {service.fields.length === 1 ? "file" : "files"}</span></div>
      {service.id === "arrears-difference" && <p className="file-guidance">Include a different date in each CSV filename, for example LoanArrears-2026-10-01.csv and LoanArrears-2026-10-16.csv. The older report is your beginning snapshot, regardless of upload order.</p>}
      <div className="upload-section-heading"><b>Source files</b><span>{selected} of {service.fields.length} selected</span></div>
      <div className="upload-list">{service.fields.map((field, index) => {
        const label = field === "targets_file" ? "Targets workbook" : field === "first_file" ? "Arrears snapshot 1" : field === "second_file" ? "Arrears snapshot 2" : field === "par_file" ? "PAR report CSV" : service.id === "repayment-report" ? "Repayment report" : service.fields.length > 1 ? "Loan report CSV" : "Source report CSV";
        const accept = field === "targets_file" ? ".xlsx,.xls" : service.id === "repayment-report" ? ".csv,.xlsx,.xls" : ".csv";
        return <label className={`dropzone ${files[field] ? "file-selected" : ""}`} key={field}>
          <span className="upload-number">{files[field] ? <CheckCircle2 size={20} /> : index + 1}</span><b>{label}</b>
          <span>{files[field] ? `${files[field].name} (${Math.max(1, Math.round(files[field].size / 1024))} KB)` : `Choose a file - ${accept.split(",").join(" or ")}`}</span>
          <input aria-label={label} type="file" accept={accept} onChange={(event) => pick(field, event)} />
        </label>;
      })}</div>
      <div className="upload-preflight"><ShieldCheck size={17} /><span>Report titles are checked automatically. Targets keep their original structure.</span></div>
      {error && <div className="error-box" role="alert">{error}</div>}
      <div className="modal-actions"><button className="secondary-btn" onClick={onClose}>Cancel</button><button className="primary-btn" onClick={submit}>Generate report <ArrowRight size={16} /></button></div>
      <p className="download-note">Your finished report downloads automatically.</p>
    </> : <div role="status" aria-live="polite"><Progress state={state} error={error} onClose={onClose} /></div>}
  </div></div>;
}
function Progress({ state, error, onClose }) { const steps = [["uploading", "Uploading files"], ["processing", "Processing report"], ["generating", "Generating report"], ["completed", "Completed"]]; const current = steps.findIndex(([id]) => id === state); return <div className="progress-view"><div className={state === "completed" ? "success-ring" : "spinner"}>{state === "completed" ? <CheckCircle2 size={34} /> : <UploadCloud size={28} />}</div><h3>{state === "completed" ? "Your report is ready" : steps[current][1]}</h3><p className="muted">{state === "completed" ? "The download has started." : "Please keep this window open while we finish your report."}</p><div className="progress-track"><div style={{ width: `${Math.max(12, ((current + 1) / steps.length) * 100)}%` }} /></div><div className="step-list">{steps.map(([id, label], index) => <div className={index <= current ? "step done" : "step"} key={id}><span>{index < current || state === "completed" ? <CheckCircle2 size={15} /> : index === current ? <span className="dot" /> : <span className="hollow" />}</span>{label}</div>)}</div>{error && <div className="error-box">{error}</div>}{state === "completed" && <button className="primary-btn full-btn" onClick={onClose}>Done</button>}</div> }
function HistoryPage({ jobs }) { return <section className="content"><p className="eyebrow">REPORT ARCHIVE</p><h1>Processing history</h1><p className="muted intro">A record of reports generated in this workspace.</p><div className="history-table"><div className="table-head"><span>Service</span><span>Output</span><span>Status</span><span>Created</span></div>{jobs.length ? jobs.map((job) => <div className="table-row" key={job.id}><b>{job.title}</b><span>{job.output}</span><span className="status-pill">Completed</span><span>{job.time}</span></div>) : <div className="table-empty"><History size={24} /><b>No processing history yet</b><span>Start a service from the workspace to see completed reports here.</span></div>}</div></section> }
function ProfilePage() { return <section className="content narrow"><p className="eyebrow">ACCOUNT</p><h1>Profile</h1><div className="settings-panel"><div className="profile-large">FT</div><div><h2>Finance team</h2><p className="muted">Workspace administrator</p></div><button className="secondary-btn">Edit profile</button></div><div className="settings-panel stacked"><label>Workspace name<input value="EasyAnalysis by Panta Development and AI" readOnly /></label><label>Workspace email<input value="finance@panta.ai" readOnly /></label></div></section> }
function SettingsPage() { return <section className="content narrow"><p className="eyebrow">WORKSPACE CONFIGURATION</p><h1>Settings</h1><div className="settings-panel stacked"><div><h2>Processing preferences</h2><p className="muted">Configure how this workspace receives processing results.</p></div><label className="toggle-row"><span><b>Automatic downloads</b><small>Download generated reports when processing completes.</small></span><input type="checkbox" defaultChecked /></label><label className="toggle-row"><span><b>Processing notifications</b><small>Show completion status in the workspace.</small></span><input type="checkbox" defaultChecked /></label></div></section> }

createRoot(document.getElementById("root")).render(<App />);



