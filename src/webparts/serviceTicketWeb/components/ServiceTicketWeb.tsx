/**
 * ServiceTicketWeb.tsx
 * -----------------------------------------------------------------------------------
 * Single-file SPFx web part component: "IT Service Desk"
 *
 * WHAT THIS NEEDS IN SHAREPOINT (create this list BEFORE using the web part):
 *
 *   List title:  IT-Service-Tickets      (change LIST_NAME below if you use another name)
 *
 *   Columns to create on that list (all "Single line of text" unless noted):
 *     - Title            (built in - stores the ticket Subject)
 *     - TicketID          Single line of text     e.g. "INC-00014"
 *     - EmployeeName       Single line of text
 *     - Email              Single line of text
 *     - Category           Single line of text     (Hardware / Software / Network)
 *     - Priority           Single line of text     (Low / Medium / Critical)
 *     - Status             Single line of text     (Open / In Progress / Resolved / Closed)
 *     - Description        Multiple lines of text  (plain text)
 *     - AssignedTo         Single line of text     (defaults to "Unassigned")
 *     - CommentsJSON       Multiple lines of text  (plain text - stores a JSON array)
 *     - TimelineJSON       Multiple lines of text  (plain text - stores a JSON array)
 *
 *   "Created" is the list's built-in date column and is used as-is.
 *
 * HOW IT WORKS:
 *   - Everything (create / read / update / delete) goes straight through the
 *     SharePoint REST API using the web part's SPHttpClient - no extra npm packages.
 *   - Deleting a ticket deletes the underlying list item too.
 *   - "Admin access" is a simple client-side gate (password: admin) that flips
 *     the UI into a workspace with a dashboard, status/assignment controls and delete.
 *     This is a UI convenience only, NOT real security - list-level permissions in
 *     SharePoint are what actually protect the data.
 *
 * SETUP:
 *   1. Create the list + columns above.
 *   2. Make sure IServiceTicketWebProps.ts exposes `context: WebPartContext`
 *      and that the web part's render() passes `context={this.context}`.
 *   3. Drop this file in unchanged at:
 *      src/webparts/serviceTicketWeb/components/ServiceTicketWeb.tsx
 *   4. gulp serve
 * -----------------------------------------------------------------------------------
 */

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

/* =====================================================================================
   PROPS
===================================================================================== */

export interface IServiceTicketWebProps {
  context: WebPartContext;
}

/* =====================================================================================
   CONSTANTS
===================================================================================== */

const LIST_NAME = 'IT-Service-Tickets';
const ADMIN_PASSWORD = 'admin';

const CATEGORIES = ['Hardware', 'Software', 'Network'];
const PRIORITIES = ['Low', 'Medium', 'Critical'];
const STATUSES = ['Open', 'In Progress', 'Resolved', 'Closed'];
const AGENTS = ['Bhardwaj, Shristi', 'Sharma, Priya', 'Kumar, Rohan', 'Verma, Ankit'];

/* =====================================================================================
   TYPES
===================================================================================== */

interface IComment {
  author: string;
  date: string;
  text: string;
}

interface ITimelineEntry {
  label: string;
  date: string;
  user: string;
  note?: string;
}

interface ISPTicketItem {
  Id: number;
  TicketID: string;
  Title: string;
  EmployeeName: string;
  Email: string;
  Category: string;
  Priority: string;
  Status: string;
  Description: string;
  AssignedTo: string;
  CommentsJSON: string;
  TimelineJSON: string;
  Created: string;
}

interface ITicket {
  id: number;
  ticketId: string;
  subject: string;
  employeeName: string;
  email: string;
  category: string;
  priority: string;
  status: string;
  description: string;
  assignedTo: string;
  comments: IComment[];
  timeline: ITimelineEntry[];
  created: string;
}

type AppMode = 'user' | 'admin';
type UserView = 'home' | 'raise' | 'ticketDetail' | 'adminPassword';
type AdminView = 'dashboard' | 'allTickets' | 'raise' | 'ticketDetail';

interface IRaiseForm {
  employeeName: string;
  email: string;
  category: string;
  priority: string;
  subject: string;
  description: string;
}

/* =====================================================================================
   HELPERS
===================================================================================== */

function padLeft(value: number, length: number): string {
  let s = String(value);
  while (s.length < length) { s = '0' + s; }
  return s;
}

function pad5(n: number): string {
  return padLeft(n, 5);
}

function formatDate(input: string | Date): string {
  const d = typeof input === 'string' ? new Date(input) : input;
  if (isNaN(d.getTime())) { return ''; }
  const day = d.getDate();
  const month = d.toLocaleString('en-US', { month: 'short' });
  const year = d.getFullYear();
  let hours = d.getHours();
  const minutes = padLeft(d.getMinutes(), 2);
  const ampm = hours >= 12 ? 'pm' : 'am';
  hours = hours % 12;
  if (hours === 0) { hours = 12; }
  return `${day} ${month} ${year}, ${hours}:${minutes} ${ampm}`;
}

function parseJsonArray<T>(raw: string | undefined | null): T[] {
  if (!raw) { return []; }
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as T[] : [];
  } catch {
    return [];
  }
}

function mapSPItemToTicket(item: ISPTicketItem): ITicket {
  return {
    id: item.Id,
    ticketId: item.TicketID || `INC-${pad5(item.Id)}`,
    subject: item.Title || '',
    employeeName: item.EmployeeName || '',
    email: item.Email || '',
    category: item.Category || '',
    priority: item.Priority || 'Medium',
    status: item.Status || 'Open',
    description: item.Description || '',
    assignedTo: item.AssignedTo || 'Unassigned',
    comments: parseJsonArray<IComment>(item.CommentsJSON),
    timeline: parseJsonArray<ITimelineEntry>(item.TimelineJSON),
    created: item.Created
  };
}

function findTicketById(list: ITicket[], id: string): ITicket | undefined {
  const target = id.toLowerCase();
  for (let i = 0; i < list.length; i++) {
    if (list[i].ticketId.toLowerCase() === target) { return list[i]; }
  }
  return undefined;
}

function nextTicketId(tickets: ITicket[]): string {
  let max = 0;
  tickets.forEach(t => {
    const m = /INC-(\d+)/.exec(t.ticketId || '');
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > max) { max = n; }
    }
  });
  return `INC-${pad5(max + 1)}`;
}

function downloadCsv(tickets: ITicket[]): void {
  const header = ['Ticket ID', 'Subject', 'Category', 'Priority', 'Status', 'Created', 'Assigned To', 'Raised By', 'Email'];
  const rows = tickets.map(t => [
    t.ticketId, t.subject, t.category, t.priority, t.status,
    formatDate(t.created), t.assignedTo, t.employeeName, t.email
  ]);
  const csv = [header, ...rows]
    .map(r => r.map(cell => `"${String(cell || '').replace(/"/g, '""')}"`).join(','))
    .join('\r\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `tickets-export-${Date.now()}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* =====================================================================================
   SHAREPOINT REST HELPERS
===================================================================================== */

async function getEntityTypeName(context: WebPartContext): Promise<string> {
  const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')?$select=ListItemEntityTypeFullName`;
  const res: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!res.ok) {
    throw new Error(`Could not find the list "${LIST_NAME}". Please create it first (see the setup notes at the top of this file).`);
  }
  const data = await res.json();
  return data.ListItemEntityTypeFullName as string;
}

async function fetchTickets(context: WebPartContext): Promise<ITicket[]> {
  const select = 'Id,TicketID,Title,EmployeeName,Email,Category,Priority,Status,Description,AssignedTo,CommentsJSON,TimelineJSON,Created';
  const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items?$select=${select}&$orderby=Created desc&$top=5000`;
  const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!res.ok) {
    throw new Error('Failed to load tickets from SharePoint.');
  }
  const data = await res.json();
  return (data.value as ISPTicketItem[]).map(mapSPItemToTicket);
}

async function createTicket(
  context: WebPartContext,
  entityType: string,
  fields: Record<string, unknown>
): Promise<ITicket> {
  const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items`;
  const body = JSON.stringify({ __metadata: { type: entityType }, ...fields });
  const res = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=verbose',
      'odata-version': ''
    },
    body
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Failed to create ticket: ${t}`);
  }
  const item = await res.json();
  return mapSPItemToTicket(item as ISPTicketItem);
}

async function addAttachmentFile(context: WebPartContext, itemId: number, file: File): Promise<void> {
  const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
  const arrayBuffer = await file.arrayBuffer();
  const res = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=nometadata',
      'odata-version': ''
    },
    body: arrayBuffer as unknown as string
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Failed to attach file: ${t}`);
  }
}

async function updateTicket(
  context: WebPartContext,
  entityType: string,
  id: number,
  fields: Record<string, unknown>
): Promise<void> {
  const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items(${id})`;
  const body = JSON.stringify({ __metadata: { type: entityType }, ...fields });
  const res = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=verbose',
      'odata-version': '',
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    },
    body
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Failed to update ticket: ${t}`);
  }
}

async function deleteTicket(context: WebPartContext, id: number): Promise<void> {
  const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items(${id})`;
  const res = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=nometadata',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'DELETE'
    }
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Failed to delete ticket: ${t}`);
  }
}

/* =====================================================================================
   STYLES  (all inline - React.CSSProperties)
===================================================================================== */

const colors = {
  navy: '#123A5B',
  navyDark: '#0E2E48',
  blue: '#1670C4',
  blueDark: '#0F5FA8',
  text: '#1F2937',
  subtext: '#6B7280',
  border: '#E5E7EB',
  bg: '#F5F6F8',
  white: '#FFFFFF',
  green: '#2F9E58',
  greenBg: '#E1F6E9',
  orange: '#B7791F',
  orangeBg: '#FBE9CF',
  red: '#C53030',
  redBg: '#FBE0E0',
  blueBadge: '#3652C8',
  blueBadgeBg: '#E3E8FD',
  grayBadge: '#6B7280',
  grayBadgeBg: '#E7E7EA'
};

const styles: { [key: string]: React.CSSProperties } = {
  page: {
    fontFamily: '"Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif',
    color: colors.text,
    background: colors.bg,
    padding: 24,
    maxWidth: 1100,
    margin: '0 auto',
    boxSizing: 'border-box'
  },
  banner: {
    background: `linear-gradient(135deg, ${colors.navy}, ${colors.navyDark})`,
    borderRadius: 10,
    padding: '24px 28px',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: 16,
    marginBottom: 20
  },
  bannerTitle: { color: colors.white, fontSize: 26, fontWeight: 700, margin: 0 },
  bannerSubtitle: { color: '#CBD9E6', fontSize: 14, marginTop: 6 },
  bannerButtons: { display: 'flex', gap: 10, flexWrap: 'wrap' },
  btnPrimary: {
    background: colors.blue, color: colors.white, border: 'none', borderRadius: 6,
    padding: '10px 18px', fontSize: 14, fontWeight: 600, cursor: 'pointer'
  },
  btnPrimaryDisabled: {
    background: '#9DBEDB', color: '#EAF1F8', border: 'none', borderRadius: 6,
    padding: '10px 18px', fontSize: 14, fontWeight: 600, cursor: 'not-allowed'
  },
  btnWhite: {
    background: colors.white, color: colors.navy, border: 'none', borderRadius: 6,
    padding: '10px 18px', fontSize: 14, fontWeight: 600, cursor: 'pointer'
  },
  btnOutline: {
    background: colors.white, color: colors.text, border: `1px solid ${colors.border}`,
    borderRadius: 6, padding: '10px 18px', fontSize: 14, fontWeight: 600, cursor: 'pointer'
  },
  btnDangerOutline: {
    background: colors.white, color: colors.red, border: `1px solid ${colors.redBg}`,
    borderRadius: 6, padding: '10px 18px', fontSize: 14, fontWeight: 600, cursor: 'pointer'
  },
  tabBar: {
    display: 'flex', gap: 4, background: colors.white, borderRadius: 8,
    padding: 6, marginBottom: 20, flexWrap: 'wrap', border: `1px solid ${colors.border}`
  },
  tab: {
    padding: '10px 16px', borderRadius: 6, fontSize: 14, fontWeight: 600,
    cursor: 'pointer', color: colors.subtext, background: 'transparent', border: 'none'
  },
  tabActive: {
    padding: '10px 16px', borderRadius: 6, fontSize: 14, fontWeight: 600,
    cursor: 'pointer', color: colors.white, background: colors.blue, border: 'none'
  },
  homeGrid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 },
  card: {
    background: colors.white, border: `1px solid ${colors.border}`, borderRadius: 10,
    padding: 24
  },
  cardTitle: { fontSize: 13, fontWeight: 700, letterSpacing: 0.5, color: colors.text, marginBottom: 10 },
  cardText: { fontSize: 14, color: colors.subtext, marginBottom: 18, lineHeight: 1.5 },
  input: {
    width: '100%', padding: '10px 12px', borderRadius: 6, border: `1px solid ${colors.border}`,
    fontSize: 14, boxSizing: 'border-box', marginBottom: 4, color: colors.text
  },
  select: {
    width: '100%', padding: '10px 12px', borderRadius: 6, border: `1px solid ${colors.border}`,
    fontSize: 14, boxSizing: 'border-box', color: colors.text, background: colors.white
  },
  textarea: {
    width: '100%', padding: '10px 12px', borderRadius: 6, border: `1px solid ${colors.border}`,
    fontSize: 14, boxSizing: 'border-box', color: colors.text, minHeight: 90, fontFamily: 'inherit',
    resize: 'vertical'
  },
  label: { display: 'block', fontSize: 13, fontWeight: 600, color: colors.text, marginBottom: 6 },
  hint: { fontSize: 12, color: colors.subtext, marginTop: 4, marginBottom: 14 },
  statGrid: {
    display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 16, marginBottom: 20
  },
  statCard: {
    background: colors.white, border: `1px solid ${colors.border}`, borderLeft: `4px solid ${colors.blue}`,
    borderRadius: 8, padding: '18px 20px'
  },
  statNumber: { fontSize: 30, fontWeight: 700, color: colors.text },
  statLabel: { fontSize: 12, fontWeight: 600, color: colors.subtext, letterSpacing: 0.5, marginTop: 4 },
  filterBar: {
    display: 'grid', gridTemplateColumns: '2fr 1fr 1fr auto', gap: 12, alignItems: 'end', marginBottom: 16
  },
  table: { width: '100%', borderCollapse: 'collapse', background: colors.white, border: `1px solid ${colors.border}`, borderRadius: 8, overflow: 'hidden' },
  th: {
    textAlign: 'left', fontSize: 11, letterSpacing: 0.5, color: colors.subtext, fontWeight: 700,
    padding: '12px 16px', borderBottom: `1px solid ${colors.border}`, background: '#FAFBFC'
  },
  td: { padding: '14px 16px', borderBottom: `1px solid ${colors.border}`, fontSize: 14, verticalAlign: 'top' },
  badge: {
    display: 'inline-flex', alignItems: 'center', gap: 6, padding: '4px 10px',
    borderRadius: 20, fontSize: 12, fontWeight: 600
  },
  dot: { width: 6, height: 6, borderRadius: '50%', display: 'inline-block' },
  detailGrid: { display: 'grid', gridTemplateColumns: '1.4fr 1fr', gap: 20, alignItems: 'start' },
  infoRowGrid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 },
  infoLabel: { fontSize: 12, color: colors.subtext, marginBottom: 4 },
  infoValue: { fontSize: 14, color: colors.text, fontWeight: 500 },
  sectionTitle: { fontSize: 13, fontWeight: 700, letterSpacing: 0.5, marginBottom: 16 },
  timelineItem: { position: 'relative', paddingLeft: 24, paddingBottom: 24 },
  timelineDot: {
    position: 'absolute', left: 0, top: 2, width: 16, height: 16, borderRadius: '50%',
    border: `2px solid ${colors.blue}`, background: colors.white, display: 'flex',
    alignItems: 'center', justifyContent: 'center', fontSize: 10, color: colors.blue
  },
  timelineLine: {
    position: 'absolute', left: 7, top: 18, width: 2, bottom: -6, background: colors.border
  },
  commentBlock: { marginBottom: 16, paddingBottom: 16, borderBottom: `1px solid ${colors.border}` },
  passwordCard: {
    background: colors.white, border: `1px solid ${colors.border}`, borderRadius: 10, padding: 24
  },
  banner_confirm: {
    background: colors.greenBg, color: colors.green, borderRadius: 8, padding: '12px 16px',
    fontSize: 14, marginBottom: 16, fontWeight: 500
  },
  banner_error: {
    background: colors.redBg, color: colors.red, borderRadius: 8, padding: '12px 16px',
    fontSize: 14, marginBottom: 16, fontWeight: 500
  },
  emptyState: { textAlign: 'center', padding: '40px 0', color: colors.subtext, fontSize: 14 }
};

function priorityBadgeStyle(priority: string): React.CSSProperties {
  switch (priority) {
    case 'Low': return { ...styles.badge, background: colors.greenBg, color: colors.green };
    case 'Critical': return { ...styles.badge, background: colors.redBg, color: colors.red };
    default: return { ...styles.badge, background: colors.orangeBg, color: colors.orange }; // Medium
  }
}

function statusBadgeStyle(status: string): React.CSSProperties {
  switch (status) {
    case 'In Progress': return { ...styles.badge, background: colors.orangeBg, color: colors.orange };
    case 'Resolved': return { ...styles.badge, background: colors.greenBg, color: colors.green };
    case 'Closed': return { ...styles.badge, background: colors.grayBadgeBg, color: colors.grayBadge };
    default: return { ...styles.badge, background: colors.blueBadgeBg, color: colors.blueBadge }; // Open
  }
}

function dotColor(style: React.CSSProperties): React.CSSProperties {
  return { ...styles.dot, background: style.color as string };
}

/* =====================================================================================
   SMALL PRESENTATIONAL COMPONENTS
===================================================================================== */

const PriorityBadge: React.FC<{ value: string }> = ({ value }) => {
  const s = priorityBadgeStyle(value);
  return <span style={s}><span style={dotColor(s)} />{value}</span>;
};

const StatusBadge: React.FC<{ value: string }> = ({ value }) => {
  const s = statusBadgeStyle(value);
  return <span style={s}><span style={dotColor(s)} />{value}</span>;
};

const StatCard: React.FC<{ number: number; label: string; color?: string }> = ({ number, label, color }) => (
  <div style={{ ...styles.statCard, borderLeftColor: color || colors.blue }}>
    <div style={styles.statNumber}>{number}</div>
    <div style={styles.statLabel}>{label}</div>
  </div>
);

/* =====================================================================================
   MAIN COMPONENT
===================================================================================== */

const emptyForm: IRaiseForm = {
  employeeName: '', email: '', category: 'Hardware', priority: 'Medium', subject: '', description: ''
};

const ServiceTicketWeb: React.FC<IServiceTicketWebProps> = ({ context }) => {
  // ---- data ----
  const [tickets, setTickets] = useState<ITicket[]>([]);
  const [entityType, setEntityType] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMsg, setErrorMsg] = useState<string>('');
  const [infoMsg, setInfoMsg] = useState<string>('');

  // ---- navigation ----
  const [appMode, setAppMode] = useState<AppMode>('user');
  const [userView, setUserView] = useState<UserView>('home');
  const [adminView, setAdminView] = useState<AdminView>('dashboard');
  const [currentTicket, setCurrentTicket] = useState<ITicket | null>(null);

  // ---- admin auth ----
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [passwordError, setPasswordError] = useState<string>('');

  // ---- raise form ----
  const [form, setForm] = useState<IRaiseForm>(emptyForm);
  const [attachedFile, setAttachedFile] = useState<File | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ---- check status (user) ----
  const [checkTicketId, setCheckTicketId] = useState<string>('');
  const [checkError, setCheckError] = useState<string>('');

  // ---- dashboard controls ----
  const [showMyTicketsOnly, setShowMyTicketsOnly] = useState<boolean>(false);
  const [search, setSearch] = useState<string>('');
  const [statusFilter, setStatusFilter] = useState<string>('All statuses');
  const [priorityFilter, setPriorityFilter] = useState<string>('All priorities');

  // ---- ticket detail controls (admin) ----
  const [moveStatusTo, setMoveStatusTo] = useState<string>('');
  const [statusNote, setStatusNote] = useState<string>('');
  const [assignAgent, setAssignAgent] = useState<string>('');
  const [newComment, setNewComment] = useState<string>('');
  const [copiedTicketId, setCopiedTicketId] = useState<boolean>(false);

  const currentUserName = context.pageContext.user.displayName;
  const currentUserEmail = context.pageContext.user.email;

  useEffect(() => {
    setForm(f => ({ ...f, employeeName: currentUserName || '', email: currentUserEmail || '' }));
    void bootstrap();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function bootstrap(): Promise<void> {
    setIsLoading(true);
    setErrorMsg('');
    try {
      const et = await getEntityTypeName(context);
      setEntityType(et);
      const data = await fetchTickets(context);
      setTickets(data);
    } catch (e) {
      setErrorMsg((e as Error).message);
    } finally {
      setIsLoading(false);
    }
  }

  async function refreshTickets(): Promise<void> {
    setIsLoading(true);
    setErrorMsg('');
    try {
      const data = await fetchTickets(context);
      setTickets(data);
    } catch (e) {
      setErrorMsg((e as Error).message);
    } finally {
      setIsLoading(false);
    }
  }

  /* ----------------------------- navigation helpers ----------------------------- */

  function goHome(): void {
    setAppMode('user');
    setUserView('home');
    setCurrentTicket(null);
    setInfoMsg('');
    setErrorMsg('');
    setCheckError('');
    setCheckTicketId('');
  }

  function openAdminPasswordModal(): void {
    setPasswordInput('');
    setPasswordError('');
    setUserView('adminPassword');
  }

  function unlockAdminView(): void {
    if (passwordInput === ADMIN_PASSWORD) {
      setAppMode('admin');
      setAdminView('dashboard');
      setUserView('home');
      setInfoMsg('');
      setErrorMsg('');
      setPasswordInput('');
      setPasswordError('');
    } else {
      setPasswordError('Incorrect password. Please try again.');
    }
  }

  function cancelAdminPassword(): void {
    setPasswordInput('');
    setPasswordError('');
    setUserView('home');
  }

  function exitAdminDashboard(): void {
    setAppMode('user');
    setUserView('home');
    setCurrentTicket(null);
  }

  function openRaiseForm(): void {
    setForm({ ...emptyForm, employeeName: currentUserName || '', email: currentUserEmail || '' });
    setAttachedFile(null);
    setInfoMsg('');
    setErrorMsg('');
    if (appMode === 'admin') { setAdminView('raise'); } else { setUserView('raise'); }
  }

  function openTicketDetail(t: ITicket): void {
    setCurrentTicket(t);
    setMoveStatusTo('');
    setStatusNote('');
    setAssignAgent('');
    setNewComment('');
    setErrorMsg('');
    setInfoMsg('');
    if (appMode === 'admin') { setAdminView('ticketDetail'); } else { setUserView('ticketDetail'); }
  }

  function backToList(): void {
    setCurrentTicket(null);
    if (appMode === 'admin') { setAdminView('dashboard'); } else { setUserView('home'); }
  }

  /* ----------------------------- form handlers ----------------------------- */

  function updateForm<K extends keyof IRaiseForm>(key: K, value: IRaiseForm[K]): void {
    setForm(prev => ({ ...prev, [key]: value }));
  }

  async function submitRaiseTicket(): Promise<void> {
    if (!form.employeeName.trim() || !form.email.trim() || !form.subject.trim()) {
      setErrorMsg('Please fill in employee name, email and subject.');
      return;
    }
    setErrorMsg('');
    setIsLoading(true);
    try {
      const ticketId = nextTicketId(tickets);
      const nowIso = new Date().toISOString();
      const timeline: ITimelineEntry[] = [
        { label: 'Open', date: nowIso, user: form.employeeName, note: 'Ticket raised' }
      ];
      const created = await createTicket(context, entityType, {
        Title: form.subject,
        TicketID: ticketId,
        EmployeeName: form.employeeName,
        Email: form.email,
        Category: form.category,
        Priority: form.priority,
        Status: 'Open',
        Description: form.description,
        AssignedTo: 'Unassigned',
        CommentsJSON: JSON.stringify([]),
        TimelineJSON: JSON.stringify(timeline)
      });
      setTickets(prev => [created, ...prev]);
      let successMsg = `Ticket ${created.ticketId} raised. The IT team can see it now.`;
      if (attachedFile) {
        try {
          await addAttachmentFile(context, created.id, attachedFile);
        } catch (attachErr) {
          successMsg = `Ticket ${created.ticketId} raised, but the attachment failed to upload: ${(attachErr as Error).message}`;
        }
      }
      setAttachedFile(null);
      setInfoMsg(successMsg);
      openTicketDetail(created);
    } catch (e) {
      setErrorMsg((e as Error).message);
    } finally {
      setIsLoading(false);
    }
  }

  function submitCheckStatus(): void {
    setCheckError('');
    if (!checkTicketId.trim()) {
      setCheckError('Please enter a ticket ID.');
      return;
    }
    const found = findTicketById(tickets, checkTicketId.trim());
    if (!found) {
      setCheckError(`No ticket found with ID "${checkTicketId.trim()}".`);
      return;
    }
    openTicketDetail(found);
  }

  /* ----------------------------- ticket detail actions (admin) ----------------------------- */

  async function handleUpdateStatus(): Promise<void> {
    if (!currentTicket || !moveStatusTo) { return; }
    setIsLoading(true);
    setErrorMsg('');
    try {
      const nowIso = new Date().toISOString();
      const entry: ITimelineEntry = {
        label: `${currentTicket.status} to ${moveStatusTo}`,
        date: nowIso,
        user: currentUserName || 'Agent',
        note: statusNote.trim() ? statusNote.trim() : undefined
      };
      const newTimeline = [...currentTicket.timeline, entry];
      await updateTicket(context, entityType, currentTicket.id, {
        Status: moveStatusTo,
        TimelineJSON: JSON.stringify(newTimeline)
      });
      const updated: ITicket = { ...currentTicket, status: moveStatusTo, timeline: newTimeline };
      setCurrentTicket(updated);
      setTickets(prev => prev.map(t => t.id === updated.id ? updated : t));
      setMoveStatusTo('');
      setStatusNote('');
    } catch (e) {
      setErrorMsg((e as Error).message);
    } finally {
      setIsLoading(false);
    }
  }

  async function handleAssignTicket(): Promise<void> {
    if (!currentTicket || !assignAgent) { return; }
    setIsLoading(true);
    setErrorMsg('');
    try {
      await updateTicket(context, entityType, currentTicket.id, { AssignedTo: assignAgent });
      const updated: ITicket = { ...currentTicket, assignedTo: assignAgent };
      setCurrentTicket(updated);
      setTickets(prev => prev.map(t => t.id === updated.id ? updated : t));
      setAssignAgent('');
    } catch (e) {
      setErrorMsg((e as Error).message);
    } finally {
      setIsLoading(false);
    }
  }

  async function handlePostComment(): Promise<void> {
    if (!currentTicket || !newComment.trim()) { return; }
    setIsLoading(true);
    setErrorMsg('');
    try {
      const entry: IComment = {
        author: currentUserName || form.employeeName || 'User',
        date: new Date().toISOString(),
        text: newComment.trim()
      };
      const newComments = [...currentTicket.comments, entry];
      await updateTicket(context, entityType, currentTicket.id, {
        CommentsJSON: JSON.stringify(newComments)
      });
      const updated: ITicket = { ...currentTicket, comments: newComments };
      setCurrentTicket(updated);
      setTickets(prev => prev.map(t => t.id === updated.id ? updated : t));
      setNewComment('');
    } catch (e) {
      setErrorMsg((e as Error).message);
    } finally {
      setIsLoading(false);
    }
  }

  function copyTicketIdToClipboard(ticketId: string): void {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(ticketId).then(() => {
        setCopiedTicketId(true);
        setTimeout(() => setCopiedTicketId(false), 1500);
      }).catch(() => { /* clipboard permission denied - silently ignore */ });
    }
  }

  async function handleDeleteTicket(): Promise<void> {
    if (!currentTicket) { return; }
    // eslint-disable-next-line no-alert
    const ok = window.confirm(`Delete ticket ${currentTicket.ticketId}? This cannot be undone.`);
    if (!ok) { return; }
    setIsLoading(true);
    setErrorMsg('');
    try {
      await deleteTicket(context, currentTicket.id);
      setTickets(prev => prev.filter(t => t.id !== currentTicket.id));
      backToList();
    } catch (e) {
      setErrorMsg((e as Error).message);
    } finally {
      setIsLoading(false);
    }
  }

  /* ----------------------------- derived data ----------------------------- */

  const scopedTickets = showMyTicketsOnly
    ? tickets.filter(t => (t.email || '').toLowerCase() === (currentUserEmail || '').toLowerCase())
    : tickets;

  const stats = {
    total: scopedTickets.length,
    open: scopedTickets.filter(t => t.status === 'Open').length,
    inProgress: scopedTickets.filter(t => t.status === 'In Progress').length,
    resolved: scopedTickets.filter(t => t.status === 'Resolved').length,
    closed: scopedTickets.filter(t => t.status === 'Closed').length
  };

  const filteredTickets = scopedTickets.filter(t => {
    const matchesSearch = !search.trim() ||
      t.ticketId.toLowerCase().indexOf(search.toLowerCase()) !== -1 ||
      t.subject.toLowerCase().indexOf(search.toLowerCase()) !== -1;
    const matchesStatus = statusFilter === 'All statuses' || t.status === statusFilter;
    const matchesPriority = priorityFilter === 'All priorities' || t.priority === priorityFilter;
    return matchesSearch && matchesStatus && matchesPriority;
  });

  function clearFilters(): void {
    setSearch('');
    setStatusFilter('All statuses');
    setPriorityFilter('All priorities');
  }

  /* ===================================================================================
     RENDER
  =================================================================================== */

  function renderMessages(): React.ReactNode {
    return (
      <>
        {infoMsg && <div style={styles.banner_confirm}>{infoMsg}</div>}
        {errorMsg && <div style={styles.banner_error}>{errorMsg}</div>}
      </>
    );
  }

  function renderAdminPasswordPrompt(): React.ReactNode {
    return (
      <div style={styles.passwordCard}>
        <label style={styles.label}>Admin password</label>
        <input
          type="password"
          style={styles.input}
          value={passwordInput}
          onChange={e => setPasswordInput(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter') { unlockAdminView(); } }}
          autoFocus
        />
        {passwordError && <div style={{ color: colors.red, fontSize: 12, margin: '6px 0' }}>{passwordError}</div>}
        <div style={{ display: 'flex', gap: 10, marginTop: 16 }}>
          <button style={styles.btnPrimary} onClick={unlockAdminView}>Unlock admin view</button>
          <button style={styles.btnOutline} onClick={cancelAdminPassword}>Cancel</button>
        </div>
      </div>
    );
  }

  function renderUserBanner(): React.ReactNode {
    return (
      <div style={styles.banner}>
        <div>
          <h1 style={styles.bannerTitle}>IT Service Desk</h1>
          <div style={styles.bannerSubtitle}>Raise a ticket, follow its progress, and reply to the IT team.</div>
        </div>
        <div style={styles.bannerButtons}>
          <button style={styles.btnPrimary} onClick={goHome}>View as user</button>
          <button style={styles.btnWhite} onClick={openAdminPasswordModal}>Admin access</button>
        </div>
      </div>
    );
  }

  function renderAdminBanner(): React.ReactNode {
    return (
      <div style={styles.banner}>
        <div>
          <h1 style={styles.bannerTitle}>Admin Dashboard</h1>
          <div style={styles.bannerSubtitle}>Manage ticket status and assignment.</div>
        </div>
        <div style={styles.bannerButtons}>
          <button style={styles.btnWhite} onClick={exitAdminDashboard}>Exit admin dashboard</button>
        </div>
      </div>
    );
  }

  function renderAdminTabs(): React.ReactNode {
    return (
      <div style={styles.tabBar}>
        <button
          style={styles.tabActive}
          onClick={() => { setCurrentTicket(null); setAdminView('dashboard'); }}
        >
          Dashboard
        </button>
        <button
          style={styles.tabActive}
          onClick={() => {
            setCurrentTicket(null);
            setAdminView('allTickets');
            setShowMyTicketsOnly(false);
            setSearch('');
            setStatusFilter('All statuses');
            setPriorityFilter('All priorities');
            void refreshTickets();
          }}
        >
          Show all tickets
        </button>
        <button style={styles.tabActive} onClick={() => downloadCsv(filteredTickets)}>Export CSV</button>
      </div>
    );
  }

  function renderHomeTiles(): React.ReactNode {
    return (
      <div style={styles.homeGrid}>
        <div style={styles.card}>
          <div style={styles.cardTitle}>RAISE A TICKET</div>
          <div style={styles.cardText}>Tell the IT team what you need help with.</div>
          <button style={styles.btnPrimary} onClick={openRaiseForm}>Raise a ticket</button>
        </div>
        <div style={styles.card}>
          <div style={styles.cardTitle}>CHECK TICKET STATUS</div>
          <div style={styles.cardText}>Enter your ticket ID to view its status, comments, and timeline.</div>
          <label style={styles.label}>Ticket ID</label>
          <input
            style={styles.input}
            placeholder="INC-00001"
            value={checkTicketId}
            onChange={e => setCheckTicketId(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') { submitCheckStatus(); } }}
          />
          {checkError && <div style={{ color: colors.red, fontSize: 12, margin: '6px 0' }}>{checkError}</div>}
          <div style={{ marginTop: 10 }}>
            <button style={styles.btnPrimary} onClick={submitCheckStatus}>Check status</button>
          </div>
        </div>
      </div>
    );
  }

  function renderRaiseForm(): React.ReactNode {
    return (
      <div style={styles.card}>
        <div style={styles.infoRowGrid}>
          <div>
            <label style={styles.label}>Employee name</label>
            <input style={styles.input} value={form.employeeName} onChange={e => updateForm('employeeName', e.target.value)} />
          </div>
          <div>
            <label style={styles.label}>Email</label>
            <input style={styles.input} value={form.email} onChange={e => updateForm('email', e.target.value)} />
          </div>
        </div>
        <div style={styles.infoRowGrid}>
          <div>
            <label style={styles.label}>Category</label>
            <select style={styles.select} value={form.category} onChange={e => updateForm('category', e.target.value)}>
              {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={styles.label}>Priority</label>
            <select style={styles.select} value={form.priority} onChange={e => updateForm('priority', e.target.value)}>
              {PRIORITIES.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
            <div style={styles.hint}>Critical means work has stopped for a team.</div>
          </div>
        </div>
        <label style={styles.label}>Subject</label>
        <input style={{ ...styles.input, marginBottom: 16 }} value={form.subject} onChange={e => updateForm('subject', e.target.value)} />
        <label style={styles.label}>Description</label>
        <textarea style={styles.textarea} value={form.description} onChange={e => updateForm('description', e.target.value)} />

        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginTop: 14, flexWrap: 'wrap' }}>
          <input
            ref={fileInputRef}
            type="file"
            style={{ display: 'none' }}
            onChange={e => setAttachedFile(e.target.files && e.target.files.length > 0 ? e.target.files[0] : null)}
          />
          <button
            type="button"
            style={{ ...styles.btnOutline, display: 'inline-flex', alignItems: 'center', gap: 8 }}
            onClick={() => fileInputRef.current && fileInputRef.current.click()}
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48" />
            </svg>
            Attach file
          </button>
          {attachedFile && (
            <span style={{
              display: 'inline-flex', alignItems: 'center', gap: 8, fontSize: 13, color: colors.text,
              background: colors.bg, border: `1px solid ${colors.border}`, borderRadius: 6, padding: '6px 10px'
            }}>
              {attachedFile.name}
              <button
                type="button"
                onClick={() => { setAttachedFile(null); if (fileInputRef.current) { fileInputRef.current.value = ''; } }}
                title="Remove attachment"
                style={{
                  background: 'transparent', border: 'none', cursor: 'pointer', padding: 0,
                  display: 'inline-flex', alignItems: 'center', color: colors.subtext
                }}
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="18" y1="6" x2="6" y2="18" />
                  <line x1="6" y1="6" x2="18" y2="18" />
                </svg>
              </button>
            </span>
          )}
        </div>

        <div style={{ display: 'flex', gap: 10, marginTop: 20 }}>
          <button style={isLoading ? styles.btnPrimaryDisabled : styles.btnPrimary} disabled={isLoading} onClick={() => void submitRaiseTicket()}>
            {isLoading ? 'Submitting...' : 'Submit ticket'}
          </button>
          <button style={styles.btnOutline} onClick={backToList}>Cancel</button>
        </div>
      </div>
    );
  }

  function renderDashboardOverview(): React.ReactNode {
    const recentTickets = scopedTickets.slice(0, 5);
    return (
      <>
        <div style={styles.statGrid}>
          <StatCard number={stats.total} label="TOTAL TICKETS" color={colors.blue} />
          <StatCard number={stats.open} label="OPEN" color={colors.blueBadge} />
          <StatCard number={stats.inProgress} label="IN PROGRESS" color={colors.orange} />
          <StatCard number={stats.resolved} label="RESOLVED" color={colors.green} />
        </div>
        <div style={{ ...styles.sectionTitle, marginTop: 4 }}>RECENT TICKETS</div>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>TICKET ID</th>
              <th style={styles.th}>SUBJECT</th>
              <th style={styles.th}>PRIORITY</th>
              <th style={styles.th}>STATUS</th>
              <th style={styles.th}>CREATED</th>
              <th style={styles.th}>ASSIGNED TO</th>
              <th style={styles.th}>ACTIONS</th>
            </tr>
          </thead>
          <tbody>
            {recentTickets.map(t => (
              <tr key={t.id}>
                <td style={{ ...styles.td, fontWeight: 600 }}>{t.ticketId}</td>
                <td style={styles.td}>
                  <div style={{ fontWeight: 600 }}>{t.subject}</div>
                  <div style={{ fontSize: 12, color: colors.subtext }}>{t.category}</div>
                </td>
                <td style={styles.td}><PriorityBadge value={t.priority} /></td>
                <td style={styles.td}><StatusBadge value={t.status} /></td>
                <td style={styles.td}>{formatDate(t.created)}</td>
                <td style={styles.td}>{t.assignedTo}</td>
                <td style={styles.td}>
                  <a href="#" style={{ color: colors.blue, fontWeight: 600, textDecoration: 'underline' }}
                    onClick={e => { e.preventDefault(); openTicketDetail(t); }}>View</a>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {recentTickets.length === 0 && !isLoading && (
          <div style={styles.emptyState}>No tickets yet.</div>
        )}
      </>
    );
  }

  function renderAllTicketsTable(): React.ReactNode {
    return (
      <>
        <div style={styles.filterBar}>
          <div>
            <label style={styles.label}>Search</label>
            <input style={styles.input} placeholder="Ticket ID or subject" value={search} onChange={e => setSearch(e.target.value)} />
          </div>
          <div>
            <label style={styles.label}>Status</label>
            <select style={styles.select} value={statusFilter} onChange={e => setStatusFilter(e.target.value)}>
              <option>All statuses</option>
              {STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label style={styles.label}>Priority</label>
            <select style={styles.select} value={priorityFilter} onChange={e => setPriorityFilter(e.target.value)}>
              <option>All priorities</option>
              {PRIORITIES.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
          <button style={styles.btnOutline} onClick={clearFilters}>Clear</button>
        </div>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>TICKET ID</th>
              <th style={styles.th}>SUBJECT</th>
              <th style={styles.th}>PRIORITY</th>
              <th style={styles.th}>STATUS</th>
              <th style={styles.th}>CREATED</th>
              <th style={styles.th}>ASSIGNED TO</th>
              <th style={styles.th}>ACTIONS</th>
            </tr>
          </thead>
          <tbody>
            {filteredTickets.map(t => (
              <tr key={t.id}>
                <td style={{ ...styles.td, fontWeight: 600 }}>{t.ticketId}</td>
                <td style={styles.td}>
                  <div style={{ fontWeight: 600 }}>{t.subject}</div>
                  <div style={{ fontSize: 12, color: colors.subtext }}>{t.category}</div>
                </td>
                <td style={styles.td}><PriorityBadge value={t.priority} /></td>
                <td style={styles.td}><StatusBadge value={t.status} /></td>
                <td style={styles.td}>{formatDate(t.created)}</td>
                <td style={styles.td}>{t.assignedTo}</td>
                <td style={styles.td}>
                  <a href="#" style={{ color: colors.blue, fontWeight: 600, textDecoration: 'underline' }}
                    onClick={e => { e.preventDefault(); openTicketDetail(t); }}>View</a>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {filteredTickets.length === 0 && !isLoading && (
          <div style={styles.emptyState}>No tickets match your filters.</div>
        )}
      </>
    );
  }

  function renderTicketDetail(): React.ReactNode {
    if (!currentTicket) { return null; }
    const t = currentTicket;
    const isAdmin = appMode === 'admin';
    return (
      <div>
        <div style={{ display: 'flex', gap: 10, marginBottom: 16 }}>
          <button style={styles.btnOutline} onClick={backToList}>Back to list</button>
          {isAdmin && (
            <button style={styles.btnDangerOutline} onClick={() => void handleDeleteTicket()}>Delete ticket</button>
          )}
        </div>
        <div style={styles.detailGrid}>
          <div>
            <div style={{ ...styles.card, marginBottom: 20 }}>
              <div style={styles.cardTitle}>TICKET INFORMATION</div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 10 }}>
                <span style={{ fontWeight: 700 }}>{t.ticketId}</span>
                <button
                  type="button"
                  onClick={() => copyTicketIdToClipboard(t.ticketId)}
                  title="Copy ticket ID"
                  style={{
                    background: 'transparent', border: 'none', cursor: 'pointer', padding: 0,
                    display: 'inline-flex', alignItems: 'center', color: colors.subtext
                  }}
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
                    <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
                  </svg>
                </button>
                {copiedTicketId && <span style={{ fontSize: 12, color: colors.green, fontWeight: 600 }}>Copied!</span>}
                <PriorityBadge value={t.priority} />
                <StatusBadge value={t.status} />
              </div>
              <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 18 }}>{t.subject}</div>
              <div style={styles.infoRowGrid}>
                <div>
                  <div style={styles.infoLabel}>Raised by</div>
                  <div style={styles.infoValue}>{t.employeeName}</div>
                </div>
                <div>
                  <div style={styles.infoLabel}>Email</div>
                  <div style={styles.infoValue}>{t.email}</div>
                </div>
                <div>
                  <div style={styles.infoLabel}>Category</div>
                  <div style={styles.infoValue}>{t.category}</div>
                </div>
                <div>
                  <div style={styles.infoLabel}>Created</div>
                  <div style={styles.infoValue}>{formatDate(t.created)}</div>
                </div>
              </div>
              <div style={{ borderTop: `1px solid ${colors.border}`, paddingTop: 14 }}>
                <div style={styles.infoLabel}>Assigned to</div>
                <div style={{ ...styles.infoValue, marginBottom: 14 }}>{t.assignedTo}</div>
                <div style={{ fontSize: 14, color: colors.text, lineHeight: 1.5 }}>{t.description}</div>
              </div>
            </div>

            <div style={styles.card}>
              <div style={styles.cardTitle}>COMMENTS ({t.comments.length})</div>
              {t.comments.length === 0 && (
                <div style={{ fontSize: 13, color: colors.subtext, marginBottom: 16 }}>No comments yet. Start the conversation.</div>
              )}
              {t.comments.map((c, i) => (
                <div key={i} style={styles.commentBlock}>
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span style={{ fontWeight: 700, fontSize: 14 }}>{c.author}</span>
                    <span style={{ fontSize: 12, color: colors.subtext }}>{formatDate(c.date)}</span>
                  </div>
                  <div style={{ fontSize: 14, marginTop: 4 }}>{c.text}</div>
                </div>
              ))}
              <label style={styles.label}>Add a comment</label>
              <textarea
                style={styles.textarea}
                placeholder="Share an update, a question, or the steps you tried."
                value={newComment}
                onChange={e => setNewComment(e.target.value)}
              />
              <div style={{ marginTop: 10 }}>
                <button
                  style={newComment.trim() && !isLoading ? styles.btnPrimary : styles.btnPrimaryDisabled}
                  disabled={!newComment.trim() || isLoading}
                  onClick={() => void handlePostComment()}
                >
                  Post comment
                </button>
              </div>
            </div>
          </div>

          <div>
            <div style={{ ...styles.card, marginBottom: 20 }}>
              <div style={styles.cardTitle}>STATUS TIMELINE</div>
              {t.timeline.map((entry, index) => (
                <div key={index} style={styles.timelineItem}>
                  <div style={styles.timelineDot}>{index === t.timeline.length - 1 ? '\u2713' : ''}</div>
                  {index < t.timeline.length - 1 && <div style={styles.timelineLine} />}
                  <div style={{ fontWeight: 700, fontSize: 14 }}>{entry.label}</div>
                  <div style={{ fontSize: 12, color: colors.subtext, marginTop: 2 }}>
                    {formatDate(entry.date)} &middot; {entry.user}
                  </div>
                  {entry.note && <div style={{ fontSize: 13, marginTop: 4 }}>{entry.note}</div>}
                </div>
              ))}
            </div>

            {isAdmin && (
              <div style={styles.card}>
                <div style={styles.cardTitle}>UPDATE TICKET</div>
                <label style={styles.label}>Move status to</label>
                <select style={{ ...styles.select, marginBottom: 10 }} value={moveStatusTo} onChange={e => setMoveStatusTo(e.target.value)}>
                  <option value="">Choose a status</option>
                  {STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
                <div style={styles.hint}>Anything typed in the comment box is saved with the status change.</div>
                <textarea
                  style={{ ...styles.textarea, minHeight: 60, marginBottom: 10 }}
                  placeholder="Optional note about this status change"
                  value={statusNote}
                  onChange={e => setStatusNote(e.target.value)}
                />
                <button
                  style={moveStatusTo && !isLoading ? styles.btnPrimary : styles.btnPrimaryDisabled}
                  disabled={!moveStatusTo || isLoading}
                  onClick={() => void handleUpdateStatus()}
                >
                  Update status
                </button>

                <div style={{ borderTop: `1px solid ${colors.border}`, marginTop: 20, paddingTop: 16 }}>
                  <label style={styles.label}>Assigned to agent</label>
                  <select style={{ ...styles.select, marginBottom: 10 }} value={assignAgent} onChange={e => setAssignAgent(e.target.value)}>
                    <option value="">Choose an agent</option>
                    {AGENTS.map(a => <option key={a} value={a}>{a}</option>)}
                  </select>
                  <button
                    style={assignAgent && !isLoading ? styles.btnPrimary : styles.btnPrimaryDisabled}
                    disabled={!assignAgent || isLoading}
                    onClick={() => void handleAssignTicket()}
                  >
                    Assign ticket
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  /* ----------------------------- top-level layout ----------------------------- */

  let content: React.ReactNode = null;

  if (appMode === 'user') {
    if (userView === 'home') { content = renderHomeTiles(); }
    else if (userView === 'raise') { content = renderRaiseForm(); }
    else if (userView === 'ticketDetail') { content = renderTicketDetail(); }
    else if (userView === 'adminPassword') { content = renderAdminPasswordPrompt(); }
  } else {
    content = (
      <>
        {renderAdminTabs()}
        {adminView === 'dashboard' && renderDashboardOverview()}
        {adminView === 'allTickets' && renderAllTicketsTable()}
        {adminView === 'raise' && renderRaiseForm()}
        {adminView === 'ticketDetail' && renderTicketDetail()}
      </>
    );
  }

  return (
    <div style={styles.page}>
      {appMode === 'user' ? renderUserBanner() : renderAdminBanner()}
      {renderMessages()}
      {isLoading && tickets.length === 0 && !errorMsg && (
        <div style={styles.emptyState}>Loading tickets...</div>
      )}
      {content}
    </div>
  );
};

export default ServiceTicketWeb;
