import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts';
import { Plus, Edit2, Trash2, GitCompare, FileText, LayoutDashboard, AlertTriangle, CheckCircle2, XCircle, ArrowLeft, Download, Save, Copy, FolderKanban, Calculator, Search, Filter, TrendingUp, TrendingDown, DollarSign, Activity, ShieldAlert, Trophy, ChevronRight, MapPin, Clock, Building2, Users, StickyNote, LogOut, Lock, Shield, Sliders, Lightbulb, Upload, Sparkles, Target, RotateCcw, X, Database, HelpCircle, Info, FileCheck, FileX, FileClock, ExternalLink, Paperclip, Calendar, ChevronDown, Eye, File } from 'lucide-react';

// ============================================================
// EXPLANATIONS DICTIONARY
// ============================================================
const EXPLAIN = {
  // KPIs
  totalProjects: 'Total number of projects you are tracking across your portfolio.',
  contract: 'Total contract value — the headline amount your client agreed to pay you, before any taxes or deductions.',
  profit: 'Final profit you keep after VAT, all deductions, partner payouts, and project costs are removed.',
  deductions: 'Combined cost of all fees, bonds, partner payouts, and custom expenses across your projects.',
  cash: 'Real cash you have in hand upfront after the bank holds back retention and you pay your fixed costs.',
  active: 'Projects currently in progress (status set to Active).',
  risky: 'Projects flagged risky — either negative cash, negative profit, or major warning signs.',
  best: 'The project projected to deliver your highest final profit. Tap to open it.',

  // Financial inputs
  vat: 'Value Added Tax. Standard rate in Kenya is 16%. The VAT amount is collected on top of net value but does not belong to you.',
  profitMarginInput: 'Your expected profit as a percentage of the net contract value (after VAT). Used to project final profit.',
  advance: 'Percentage of contract value paid to you upfront when the project starts. Higher is better for cash flow.',
  bankRetention: 'Percentage of your advance that the bank holds as security. You only get the rest as working cash.',

  // Costs
  mobilisation: 'Cost to start the project — moving equipment, setting up site, hiring initial team.',
  performanceBond: 'Insurance cost paid to a bank or insurer guaranteeing you will deliver. Usually 1-3% of contract.',
  bankGuarantee: 'Bank fee for issuing guarantees demanded by client (e.g., advance guarantee, retention guarantee).',
  insurance: 'Construction insurance, contractors all-risk, third-party liability. Typically 0.5-2% of contract.',
  tender: 'Cost paid to bid for the project — application fees, document purchase, presentation costs.',
  legal: 'Lawyer fees, document drafting, contract review.',
  commission: 'Payments to brokers, agents, or referrers who helped you secure the project.',
  customDeductions: 'Project-specific costs you add manually — anything that does not fit standard categories.',
  partners: 'Partners or shareholders sharing profits. Each gets a percentage of final profit plus optional fixed payouts.',

  // Calculated metrics
  netAfterVat: 'Contract value minus VAT. This is the net amount your business is actually paid.',
  expectedProfit: 'Profit before any deductions are taken — your gross expected return on the project.',
  advanceCash: 'The actual cash you receive upfront after the bank takes its retention. Your starting working capital.',
  realCash: 'Cash left in your hand upfront after deducting all fixed costs, bonds, fees, and partner payouts. Negative means you fund the project out of pocket.',
  finalProfit: 'Your actual take-home profit after every deduction. This is what you walk away with.',
  effectiveMargin: 'Final profit as a percentage of contract value. Industry-comparable measure of profitability.',
  capitalEfficiency: 'Profit divided by total deductions. Multiplier showing how much profit you generate per shilling of cost. Above 3x is strong.',
  riskLevel: 'Auto-calculated rating: SAFE (healthy), CAUTION (some warning signs), RISKY (negative cash or profit).',

  // Compare
  upfrontCash: 'How much working capital you receive upfront after bank retention.',

  // Project list
  status: 'Where the project is in its lifecycle: Planning, Active, Completed, or Cancelled.',
  client: 'The company or party paying for the project.',
  location: 'Where the project is being executed.',
  type: 'Category of work — Construction, Supply, Land Subdivision, etc.'
};

// ============================================================
// CALCULATION ENGINE
// ============================================================
const calcProject = (p) => {
  const contract = +p.contractValue || 0;
  const vatPct = +p.vatPct || 0;
  const profitPct = +p.profitPct || 0;
  const advancePct = +p.advancePct || 0;
  const bankRetentionPct = +p.bankRetentionPct || 0;

  const vatAmount = contract * (vatPct / 100);
  const netAfterVat = contract - vatAmount;
  const expectedProfit = netAfterVat * (profitPct / 100);
  const advanceAmount = contract * (advancePct / 100);
  const bankRetained = advanceAmount * (bankRetentionPct / 100);
  const advanceCashAvailable = advanceAmount - bankRetained;

  const fixedCosts =
    (+p.mobilisationCost || 0) + (+p.performanceBondCost || 0) +
    (+p.bankGuaranteeCost || 0) + (+p.insuranceCost || 0) +
    (+p.tenderCost || 0) + (+p.legalCost || 0) + (+p.commissionCost || 0);

  const customTotal = (p.customDeductions || []).reduce((s, d) => s + (+d.amount || 0), 0);
  const partnerTotal = (p.partners || []).reduce((s, pt) => s + (+pt.payout || 0), 0);

  const totalDeductions = fixedCosts + customTotal + partnerTotal;
  const realCashAfterDeductions = advanceCashAvailable - totalDeductions;
  const projectedFinalProfit = expectedProfit - totalDeductions;

  let riskLevel = 'safe';
  let riskReasons = [];
  if (realCashAfterDeductions < 0) { riskLevel = 'risky'; riskReasons.push('Upfront cash negative — you will fund the project out of pocket'); }
  if (projectedFinalProfit < 0) { riskLevel = 'risky'; riskReasons.push('Final profit is negative after deductions'); }
  if (projectedFinalProfit >= 0 && projectedFinalProfit < expectedProfit * 0.3) {
    if (riskLevel !== 'risky') riskLevel = 'caution';
    riskReasons.push('Deductions consume more than 70% of expected profit');
  }
  if (advancePct < 10 && advancePct > 0) {
    if (riskLevel === 'safe') riskLevel = 'caution';
    riskReasons.push('Low advance payment — high working capital needed');
  }
  if (totalDeductions > contract * 0.15) {
    if (riskLevel === 'safe') riskLevel = 'caution';
    riskReasons.push('Deductions exceed 15% of contract value');
  }
  if (riskReasons.length === 0) riskReasons.push('All indicators within healthy range');

  const partnersWithShare = (p.partners || []).map(pt => ({
    ...pt, profitShare: projectedFinalProfit * ((+pt.sharePct || 0) / 100)
  }));

  const profitMargin = contract > 0 ? (projectedFinalProfit / contract) * 100 : 0;
  const capitalEfficiency = totalDeductions > 0 ? projectedFinalProfit / totalDeductions : null;

  return {
    contract, vatAmount, netAfterVat, expectedProfit,
    advanceAmount, bankRetained, advanceCashAvailable,
    fixedCosts, customTotal, partnerTotal, totalDeductions,
    realCashAfterDeductions, projectedFinalProfit,
    riskLevel, riskReasons, partnersWithShare,
    profitMargin, capitalEfficiency
  };
};

// ============================================================
// CASH FLOW TIMELINE — month-by-month projection
// ============================================================
const buildCashflow = (p, c) => {
  const months = parseInt(String(p.duration || '12').match(/\d+/)?.[0] || 12, 10);
  const periods = [];

  // Month 0 — upfront outlays
  const m0Out = c.totalDeductions; // sunk + locked at start
  const m0In = c.advanceAmount;     // advance received
  const m0BankRet = c.bankRetained; // held by bank
  let cumulative = m0In - m0BankRet - m0Out;
  periods.push({
    month: 0, label: 'Start',
    inflow: m0In - m0BankRet,
    outflow: m0Out,
    netCash: m0In - m0BankRet - m0Out,
    cumulative,
    note: 'Advance net of retention, less upfront costs'
  });

  // Revenue billed in even monthly draws (simple straight-line model)
  const contractMinusAdvance = c.contract - c.advanceAmount; // remaining billable
  const monthlyBilling = months > 0 ? contractMinusAdvance / months : 0;
  // Operating cost assumption: 85% of contract spent across months (rough construction COGS)
  const totalOpCost = c.contract * 0.85 - c.totalDeductions;
  const monthlyOpCost = months > 0 ? Math.max(totalOpCost / months, 0) : 0;

  for (let m = 1; m <= months; m++) {
    const inflow = monthlyBilling;
    const outflow = monthlyOpCost;
    const net = inflow - outflow;
    cumulative += net;
    periods.push({
      month: m,
      label: `M${m}`,
      inflow,
      outflow,
      netCash: net,
      cumulative,
      note: ''
    });
  }

  // Final month — retention release
  const lastIdx = periods.length - 1;
  periods[lastIdx].inflow += m0BankRet;
  periods[lastIdx].netCash += m0BankRet;
  periods[lastIdx].cumulative += m0BankRet;
  periods[lastIdx].note = 'Includes bank retention release';

  const minCash = Math.min(...periods.map(x => x.cumulative));
  const breakEvenMonth = periods.find(x => x.cumulative >= 0)?.month;
  const peakNegative = periods.reduce((a, b) => b.cumulative < a.cumulative ? b : a);

  return { periods, months, minCash, breakEvenMonth, peakNegative, monthlyBilling, monthlyOpCost };
};

const fmt = (n, opts = {}) => {  if (n === null || n === undefined || isNaN(n)) return '—';
  const { currency = 'KES', decimals = 0 } = opts;
  return `${currency} ${Number(n).toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals })}`;
};

const fmtShort = (n, currency = 'KES') => {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1_000_000) return `${sign}${currency} ${(abs / 1_000_000).toFixed(2)}M`;
  if (abs >= 1_000) return `${sign}${currency} ${(abs / 1_000).toFixed(1)}K`;
  return `${sign}${currency} ${Math.round(abs)}`;
};

const fmtPct = (n) => `${Number(n || 0).toFixed(1)}%`;

// ============================================================
// DOCUMENT TEMPLATES — Kenya construction/contracting standard set
// ============================================================
const DOCUMENT_TEMPLATES = [
  // Contract & Award
  { id: 'contract', label: 'Signed Contract', category: 'Contract & Award', critical: true },
  { id: 'loa', label: 'Letter of Award', category: 'Contract & Award', critical: true },
  { id: 'boq', label: 'Bill of Quantities (BOQ)', category: 'Contract & Award', critical: true },
  { id: 'tender', label: 'Tender Document', category: 'Contract & Award' },
  { id: 'drawings', label: 'Technical Drawings / Plans', category: 'Contract & Award' },
  // Bonds & Guarantees
  { id: 'bidBond', label: 'Bid Bond', category: 'Bonds & Guarantees', hasExpiry: true },
  { id: 'pb', label: 'Performance Bond (PB)', category: 'Bonds & Guarantees', critical: true, hasExpiry: true },
  { id: 'apg', label: 'Advance Payment Guarantee (APG)', category: 'Bonds & Guarantees', critical: true, hasExpiry: true },
  { id: 'retentionBond', label: 'Retention Bond', category: 'Bonds & Guarantees', hasExpiry: true },
  // Insurance
  { id: 'allRisk', label: 'All-Risk Insurance', category: 'Insurance', critical: true, hasExpiry: true },
  { id: 'profIndemnity', label: 'Professional Indemnity', category: 'Insurance', hasExpiry: true },
  { id: 'publicLiability', label: 'Public Liability', category: 'Insurance', hasExpiry: true },
  { id: 'workComp', label: "Workmen's Compensation (WIBA)", category: 'Insurance', critical: true, hasExpiry: true },
  // Kenya Compliance
  { id: 'nema', label: 'NEMA Environmental Approval', category: 'Compliance', hasExpiry: true },
  { id: 'nca', label: 'NCA Registration', category: 'Compliance', hasExpiry: true },
  { id: 'buildingPermit', label: 'Building / Construction Permit', category: 'Compliance' },
  { id: 'taxCert', label: 'Tax Compliance Certificate', category: 'Compliance', hasExpiry: true },
  { id: 'kraPin', label: 'KRA PIN Certificate', category: 'Compliance' },
  { id: 'vatCert', label: 'VAT Registration Certificate', category: 'Compliance' },
  // Financial
  { id: 'invoices', label: 'Invoices', category: 'Financial' },
  { id: 'receipts', label: 'Receipts', category: 'Financial' },
  { id: 'bankStatement', label: 'Bank Statements', category: 'Financial' },
  { id: 'quotations', label: 'Supplier Quotations', category: 'Financial' },
  { id: 'po', label: 'Purchase Orders', category: 'Financial' },
  // Partner & Legal
  { id: 'partnership', label: 'Partnership / JV Agreement', category: 'Partner & Legal' },
  { id: 'mou', label: 'Memorandum of Understanding (MoU)', category: 'Partner & Legal' },
  { id: 'nda', label: 'Non-Disclosure Agreement (NDA)', category: 'Partner & Legal' },
  { id: 'poa', label: 'Power of Attorney', category: 'Partner & Legal' },
  // Site & Reports
  { id: 'sitePhotos', label: 'Site Photos', category: 'Site & Reports' },
  { id: 'progressReport', label: 'Progress Reports', category: 'Site & Reports' },
  { id: 'inspectionCert', label: 'Inspection Certificates', category: 'Site & Reports' },
];

const DOC_CATEGORIES = ['Contract & Award', 'Bonds & Guarantees', 'Insurance', 'Compliance', 'Financial', 'Partner & Legal', 'Site & Reports', 'Custom'];

const getDocTemplate = (templateId) => DOCUMENT_TEMPLATES.find(t => t.id === templateId);

// Compute document compliance score for a project
const computeDocStatus = (project) => {
  const docs = project.documents || [];
  const criticalTemplates = DOCUMENT_TEMPLATES.filter(t => t.critical);
  const criticalReceived = criticalTemplates.filter(t => {
    const d = docs.find(x => x.templateId === t.id);
    return d && d.status === 'received';
  });
  const missing = criticalTemplates.filter(t => {
    const d = docs.find(x => x.templateId === t.id);
    return !d || d.status === 'pending' || d.status === 'required';
  });
  const expiringSoon = docs.filter(d => {
    if (!d.expiryDate || d.status !== 'received') return false;
    const days = Math.ceil((new Date(d.expiryDate) - new Date()) / (1000 * 60 * 60 * 24));
    return days >= 0 && days <= 30;
  });
  const expired = docs.filter(d => {
    if (!d.expiryDate || d.status !== 'received') return false;
    return new Date(d.expiryDate) < new Date();
  });
  const compliancePct = criticalTemplates.length > 0 ? (criticalReceived.length / criticalTemplates.length) * 100 : 100;
  return { criticalTemplates, criticalReceived, missing, expiringSoon, expired, compliancePct, total: docs.length };
};


// ============================================================
// PORTFOLIO ANALYTICS
// ============================================================
const computeHealthScore = (p, c) => {
  let score = 50;
  const factors = [];

  if (c.realCashAfterDeductions > 0) { score += 15; factors.push({ label: 'Cash positive upfront', delta: 15, good: true }); }
  else { score -= 5; factors.push({ label: 'Negative cash upfront', delta: -5, good: false }); }

  if (c.projectedFinalProfit > 0) { score += 15; factors.push({ label: 'Profit positive', delta: 15, good: true }); }
  else { score -= 15; factors.push({ label: 'Profit negative', delta: -15, good: false }); }

  if (c.profitMargin > 20) { score += 15; factors.push({ label: 'Strong margin (>20%)', delta: 15, good: true }); }
  else if (c.profitMargin > 10) { score += 10; factors.push({ label: 'Healthy margin (>10%)', delta: 10, good: true }); }
  else if (c.profitMargin > 0) { score += 3; factors.push({ label: 'Thin margin', delta: 3, good: false }); }

  if (c.capitalEfficiency && c.capitalEfficiency > 3) { score += 10; factors.push({ label: 'Efficient capital use (>3x)', delta: 10, good: true }); }
  else if (c.capitalEfficiency && c.capitalEfficiency > 1) { score += 5; factors.push({ label: 'Acceptable capital use', delta: 5, good: true }); }

  const isBigContract = c.contract > 100_000_000;
  if (+p.performanceBondCost > 0 || +p.bankGuaranteeCost > 0) { score += 5; factors.push({ label: 'Has bond/guarantee', delta: 5, good: true }); }
  else if (isBigContract) { score -= 10; factors.push({ label: 'Missing bond on large contract', delta: -10, good: false }); }

  if (+p.insuranceCost > 0) { score += 5; factors.push({ label: 'Insurance recorded', delta: 5, good: true }); }
  else if (isBigContract) { score -= 5; factors.push({ label: 'Missing insurance', delta: -5, good: false }); }

  if (c.totalDeductions > c.contract * 0.15) { score -= 10; factors.push({ label: 'High deductions (>15% of contract)', delta: -10, good: false }); }

  return { score: Math.max(0, Math.min(100, score)), factors };
};

const applyExpenseFactor = (p, factor) => ({
  ...p,
  mobilisationCost: (+p.mobilisationCost || 0) * factor,
  performanceBondCost: (+p.performanceBondCost || 0) * factor,
  bankGuaranteeCost: (+p.bankGuaranteeCost || 0) * factor,
  insuranceCost: (+p.insuranceCost || 0) * factor,
  tenderCost: (+p.tenderCost || 0) * factor,
  legalCost: (+p.legalCost || 0) * factor,
  commissionCost: (+p.commissionCost || 0) * factor,
  customDeductions: (p.customDeductions || []).map(d => ({ ...d, amount: (+d.amount || 0) * factor }))
});

const computeStressTests = (project) => {
  const baseline = calcProject(project);
  const tests = [
    { label: 'Margin drops 5pp', sub: 'Profit margin reduced by 5 percentage points', mod: { profitPct: Math.max(0, +project.profitPct - 5) } },
    { label: 'Expenses +25%', sub: 'All deductions increased by 25%', expenseFactor: 1.25 },
    { label: 'Advance halved', sub: 'Upfront advance payment cut in half', mod: { advancePct: (+project.advancePct || 0) / 2 } },
    { label: 'Worst case combined', sub: 'Margin -3pp, expenses +20%, advance -25%', mod: { profitPct: Math.max(0, +project.profitPct - 3), advancePct: (+project.advancePct || 0) * 0.75 }, expenseFactor: 1.2 }
  ];

  return tests.map(t => {
    let modified = { ...project, ...(t.mod || {}) };
    if (t.expenseFactor) modified = applyExpenseFactor(modified, t.expenseFactor);
    const calc = calcProject(modified);
    return {
      label: t.label,
      sub: t.sub,
      profit: calc.projectedFinalProfit,
      cash: calc.realCashAfterDeductions,
      risk: calc.riskLevel,
      profitDelta: calc.projectedFinalProfit - baseline.projectedFinalProfit,
      cashDelta: calc.realCashAfterDeductions - baseline.realCashAfterDeductions
    };
  });
};

const computeRecommendations = (calcs) => {
  const recs = [];
  calcs.forEach(({ p, c }) => {
    // Document compliance recommendations
    const ds = computeDocStatus(p);
    ds.missing.forEach(t => {
      recs.push({ priority: t.id === 'pb' || t.id === 'apg' ? 'high' : 'medium', icon: 'FileText', title: `${t.label} missing on ${p.name}`, body: `Critical document not yet recorded. Upload to Drive and link it to this project.`, projectId: p.id });
    });
    ds.expired.forEach(d => {
      const t = d.templateId ? getDocTemplate(d.templateId) : null;
      const label = t?.label || d.customLabel || 'Document';
      recs.push({ priority: 'high', icon: 'AlertTriangle', title: `${label} EXPIRED on ${p.name}`, body: `Renew immediately to maintain compliance.`, projectId: p.id });
    });
    ds.expiringSoon.forEach(d => {
      const t = d.templateId ? getDocTemplate(d.templateId) : null;
      const label = t?.label || d.customLabel || 'Document';
      const days = Math.ceil((new Date(d.expiryDate) - new Date()) / (1000 * 60 * 60 * 24));
      recs.push({ priority: days <= 7 ? 'high' : 'medium', icon: 'Clock', title: `${label} expires in ${days} days on ${p.name}`, body: `Schedule renewal to avoid lapsed coverage.`, projectId: p.id });
    });

    if (+p.advancePct < 15 && +p.advancePct > 0 && c.contract > 50_000_000) {
      recs.push({ priority: 'high', icon: 'DollarSign', title: `Negotiate higher advance on ${p.name}`, body: `Currently ${fmtPct(p.advancePct)} on ${fmtShort(c.contract, p.currency)} contract. Even 5% more = ${fmtShort(c.contract * 0.05, p.currency)} extra upfront.`, projectId: p.id });
    }
    if (c.contract > 100_000_000 && +p.performanceBondCost === 0 && +p.bankGuaranteeCost === 0) {
      recs.push({ priority: 'high', icon: 'Shield', title: `Add bond/guarantee cost for ${p.name}`, body: 'Large contracts typically require performance bonds (1-3% of contract). Track this or face surprise costs.', projectId: p.id });
    }
    if (c.contract > 100_000_000 && +p.insuranceCost === 0) {
      recs.push({ priority: 'medium', icon: 'Shield', title: `Track insurance cost for ${p.name}`, body: 'No insurance recorded. Construction insurance typically 0.5-2% of contract value.', projectId: p.id });
    }
    if (c.profitMargin > 30) {
      recs.push({ priority: 'medium', icon: 'AlertTriangle', title: `Pressure-test margin on ${p.name}`, body: `${c.profitMargin.toFixed(1)}% effective margin is unusually high. Verify before committing.`, projectId: p.id });
    }
    if (c.realCashAfterDeductions < 0) {
      recs.push({ priority: 'high', icon: 'AlertTriangle', title: `Cash gap on ${p.name}`, body: `Need ${fmtShort(Math.abs(c.realCashAfterDeductions), p.currency)} extra working capital. Consider partner, financing, or push for higher advance.`, projectId: p.id });
    }
    const partnerSharePct = c.expectedProfit > 0 ? (c.partnerTotal / c.expectedProfit) * 100 : 0;
    if (partnerSharePct > 40 && c.partnerTotal > 0) {
      recs.push({ priority: 'medium', icon: 'Users', title: `High partner exposure on ${p.name}`, body: `${partnerSharePct.toFixed(0)}% of expected profit going to partners. Re-verify share terms.`, projectId: p.id });
    }
  });

  if (calcs.length >= 2) {
    const totalContract = calcs.reduce((s, x) => s + x.c.contract, 0);
    const sorted = [...calcs].sort((a, b) => b.c.contract - a.c.contract);
    if (sorted[0] && totalContract > 0 && sorted[0].c.contract / totalContract > 0.7) {
      recs.push({ priority: 'medium', icon: 'ShieldAlert', title: 'Portfolio is over-concentrated', body: `${((sorted[0].c.contract / totalContract) * 100).toFixed(0)}% of portfolio value in ${sorted[0].p.name}. Consider diversifying.`, projectId: sorted[0].p.id });
    }
    const planning = calcs.filter(x => x.p.status === 'planning');
    const active = calcs.filter(x => x.p.status === 'active');
    if (planning.length >= 3 && active.length === 0) {
      recs.push({ priority: 'medium', icon: 'Activity', title: 'Planning bottleneck', body: `${planning.length} projects in planning, none active. Pick the strongest and move it forward.` });
    }
  }

  return recs.sort((a, b) => ({ high: 0, medium: 1, low: 2 }[a.priority] - { high: 0, medium: 1, low: 2 }[b.priority]));
};

const blankProject = () => ({
  id: `p_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
  name: '', client: '', location: '', type: '',
  startDate: '', duration: '', status: 'planning', currency: 'KES',
  contractValue: 0, vatPct: 16, profitPct: 15, advancePct: 20, bankRetentionPct: 10,
  mobilisationCost: 0, performanceBondCost: 0, bankGuaranteeCost: 0,
  insuranceCost: 0, tenderCost: 0, legalCost: 0, commissionCost: 0,
  customDeductions: [], partners: [], notes: '', documents: [],
  createdAt: new Date().toISOString()
});

// ============================================================
// STORAGE
// ============================================================
const STORAGE_KEY = 'rays_projects_v7';
const SEED_PROJECTS = [
  {
    id: 'p_kanyakwar_phase3',
    name: 'Upper Kanyakwar Phase 3 AHP',
    client: 'Ministry of Lands / GoK',
    location: 'Kanyakwar, Kisumu County',
    type: 'Affordable Housing / Construction',
    startDate: '', duration: '24 months', status: 'planning', currency: 'KES',
    contractValue: 2297404001.19,
    vatPct: 16,
    profitPct: 17.86,        // 17.86% × net-after-VAT = 344.6M (= 15% gross of contract)
    advancePct: 20,          // Two APGs of 10% each = 20% of contract
    bankRetentionPct: 20,    // 20% cash margin held at CIB against both APGs
    mobilisationCost: 17727100.18,        // Pre-mobilization fees Y1 (CAR + WIBA-B + APG1 fee + CIC bond 1)
    performanceBondCost: 0,                // APGs serve as the bonds
    bankGuaranteeCost: 12061371.00,        // CIB APG1 fee 6,892,212 + APG2 fee 5,169,159
    insuranceCost: 7534864.80,             // CAR 2,596,250 + WIBA-B × 2yrs 4,938,614.80
    tenderCost: 0, legalCost: 0, commissionCost: 0,
    customDeductions: [
      { id: 'cd_cic_fee', label: 'CIC counter-guarantee fee (both APGs)', amount: 5538661.56 },
      { id: 'cd_facility_interest', label: 'Additional facility interest (14% × 50M × 12mo)', amount: 7000000 },
      { id: 'cd_cic_collateral', label: 'CIC collateral lock — RECOVERABLE cash', amount: 73516928.04 }
    ],
    partners: [],
    documents: [
      { templateId: 'contract', status: 'pending' },
      { templateId: 'loa', status: 'pending' },
      { templateId: 'boq', status: 'received', notes: 'Grand Sum: KES 2,297,404,001.19' },
      { templateId: 'apg', status: 'pending', notes: 'APG 1: 229,740,400.12 — 10% of contract' },
      { id: 'cust_apg2', templateId: null, customLabel: 'Advance Payment Guarantee 2 (APG 2)', category: 'Bonds & Guarantees', status: 'pending', notes: 'APG 2: 229,740,400.12 — 10% of contract' },
      { templateId: 'allRisk', status: 'pending', notes: 'CAR Insurance — KES 2,596,250 total premium' },
      { templateId: 'workComp', status: 'pending', notes: 'WIBA + EL (Scenario B × 2yrs): KES 4,938,614.80' },
      { id: 'cust_cic', templateId: null, customLabel: 'CIC Counter-Guarantee Bond', category: 'Bonds & Guarantees', status: 'pending', notes: 'Counter-guarantee covering both APGs. Cash collateral: 73,516,928.04 (or 40% assets = 147,033,856.08).' },
      { id: 'cust_facility', templateId: null, customLabel: 'Additional Facility Letter (50M @ 14%)', category: 'Financial', status: 'pending' },
      { templateId: 'kraPin', status: 'pending' },
      { templateId: 'vatCert', status: 'pending' },
      { templateId: 'nca', status: 'pending', notes: 'NCA registration valid for the contract period' },
      { templateId: 'nema', status: 'pending' },
      { templateId: 'taxCert', status: 'pending' }
    ],
    notes: `PROJECT BASICS
• Contract Sum: KES 2,297,404,001.19
• Duration: 24 months
• Location: Kisumu, Kenya
• Client: Ministry of Lands / GoK (low default risk)

FINANCING STRUCTURE
• APG 1 advance: +229,740,400.12 (10% of contract)
• APG 2 advance: +229,740,400.12 (10% of contract)
• Additional facility: +50,000,000.00 @ 14% × 12mo
• Total finance available: 509,480,800.24

CAPITAL REQUIRED (CASH LOCKED / DEPLOYED)
• Cash margin APG 1: 45,948,080.02
• Cash margin APG 2: 45,948,080.02
• CIC collateral (20% cash on both bonds): 73,516,928.04
  └─ Or 40% in assets = 147,033,856.08
• Pre-mobilization fees Y1: 17,727,100.18
  └─ CAR + WIBA-B + APG1 fee + CIC bond 1
• TOTAL UPFRONT CASH: 183,140,188.26
• If asset collateral at CIC: 109,623,260.22

NON-RECOVERABLE COST OF CAPITAL (24 months)
• CAR Insurance: 2,596,250.00
• WIBA + EL (Scenario B × 2yrs): 4,938,614.80
• CIB APG 1 fee: 6,892,212.00
• CIB APG 2 fee: 5,169,159.00
• CIC counter-guarantee (both): 5,538,661.56
• Additional facility interest (14% × 12mo × 50M): 7,000,000.00
• TOTAL COST OF CAPITAL: 32,134,897.36

KEY RATIOS
• Cost of capital / Contract: 1.40%
• Cost of capital / Total finance raised: 6.31%
• Cash margin / Contract: 4.00%
• Total cash locked / Contract: 7.97%

REVENUE / MARGIN ASSUMPTIONS
• Kenyan AHP gross margin range: 12-18%
• At 15% gross: 344,610,600 gross profit
• Less financing cost: 312,475,703 net before overheads & tax
• ROI on capital deployed (cash basis): ~170% over 24 months @ 15% gross

NOTE: Calculator's Profit % = 17.86% applied to net-after-VAT to yield 344.6M = 15% gross on contract.
The CIC collateral (73.5M) is shown as a deduction so "Real Cash Available" reflects deployable cash; it is RECOVERABLE at project end.`,
    createdAt: new Date().toISOString()
  }
];
const loadProjects = async () => {
  try {
    const r = await window.storage.get(STORAGE_KEY);
    return r ? JSON.parse(r.value) : SEED_PROJECTS;
  } catch {
    try { await window.storage.set(STORAGE_KEY, JSON.stringify(SEED_PROJECTS)); } catch {}
    return SEED_PROJECTS;
  }
};
const saveProjects = async (projects) => {
  try { await window.storage.set(STORAGE_KEY, JSON.stringify(projects)); }
  catch (e) { console.error('Storage failed', e); }
};

// ============================================================
// AUTH STORAGE
// ============================================================
const AUTH_KEY = 'rays_auth_v1';

const hashPassword = async (pw) => {
  const enc = new TextEncoder();
  const data = enc.encode(pw);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  return Array.from(new Uint8Array(hashBuffer)).map(b => b.toString(16).padStart(2, '0')).join('');
};

const loadAuthHash = async () => {
  try {
    const r = await window.storage.get(AUTH_KEY);
    return r ? r.value : null;
  } catch { return null; }
};

const saveAuthHash = async (hash) => {
  try { await window.storage.set(AUTH_KEY, hash); }
  catch (e) { console.error('Auth save failed', e); }
};

const clearAuthHash = async () => {
  try { await window.storage.delete(AUTH_KEY); } catch {}
};

// ============================================================
// DOCUMENT FILE STORAGE
// Each uploaded file gets its own storage key so projects array stays small
// ============================================================
const MAX_FILE_BYTES = 4 * 1024 * 1024; // 4MB hard limit per file
const docFileKey = (projectId, docKey) => `rays_docfile_${projectId}_${docKey}`;

const formatBytes = (bytes) => {
  if (!bytes) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let i = 0;
  let n = bytes;
  while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
  return `${n.toFixed(n >= 10 || i === 0 ? 0 : 1)} ${units[i]}`;
};

const uploadDocFile = async (projectId, docKey, file) => {
  if (file.size > MAX_FILE_BYTES) {
    throw new Error(`File is ${formatBytes(file.size)}. Max allowed: ${formatBytes(MAX_FILE_BYTES)}. For larger files, upload to Google Drive and paste the link instead.`);
  }
  const dataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('Could not read file'));
    reader.readAsDataURL(file);
  });
  await window.storage.set(docFileKey(projectId, docKey), dataUrl);
  return { fileName: file.name, fileSize: file.size, fileType: file.type, uploadedAt: new Date().toISOString() };
};

const downloadDocFile = async (projectId, docKey, fileName) => {
  try {
    const r = await window.storage.get(docFileKey(projectId, docKey));
    if (!r || !r.value) throw new Error('File not found in storage');
    const a = document.createElement('a');
    a.href = r.value;
    a.download = fileName || 'document';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (e) {
    console.error('Download failed:', e);
    alert('Could not download file: ' + e.message);
  }
};

const previewDocFile = async (projectId, docKey) => {
  try {
    const r = await window.storage.get(docFileKey(projectId, docKey));
    if (!r || !r.value) throw new Error('File not found');
    const w = window.open();
    if (w) {
      w.document.write(`<iframe src="${r.value}" style="width:100%;height:100vh;border:0;"></iframe>`);
    }
  } catch (e) {
    alert('Could not preview file: ' + e.message);
  }
};

const deleteDocFile = async (projectId, docKey) => {
  try { await window.storage.delete(docFileKey(projectId, docKey)); } catch {}
};

// ============================================================
// REPORT GENERATORS — Executive PDF + Excel exports
// ============================================================

const REPORT_CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Inter', -apple-system, sans-serif; color: #0f172a; line-height: 1.5; padding: 40px 48px; max-width: 920px; margin: 0 auto; background: white; }
  .brand-bar { display: flex; justify-content: space-between; align-items: center; padding-bottom: 16px; margin-bottom: 32px; border-bottom: 3px solid #0f172a; }
  .brand-name { font-family: 'Manrope', sans-serif; font-weight: 800; font-size: 14px; letter-spacing: 0.18em; color: #0f172a; }
  .brand-tag { font-size: 10px; color: #64748b; text-transform: uppercase; letter-spacing: 0.1em; font-weight: 700; }
  .doc-title { font-family: 'Manrope', sans-serif; font-weight: 800; font-size: 36px; line-height: 1.1; margin-bottom: 6px; letter-spacing: -0.02em; color: #0f172a; }
  .doc-subtitle { font-size: 14px; color: #64748b; margin-bottom: 8px; }
  .doc-meta { display: flex; gap: 16px; flex-wrap: wrap; font-size: 12px; color: #475569; padding: 12px 0; border-bottom: 1px solid #e2e8f0; margin-bottom: 32px; }
  .doc-meta strong { color: #0f172a; font-weight: 700; margin-right: 6px; }
  h2 { font-family: 'Manrope', sans-serif; font-weight: 700; font-size: 18px; margin: 32px 0 14px; padding-bottom: 8px; border-bottom: 2px solid #0f172a; letter-spacing: -0.01em; }
  h3 { font-family: 'Manrope', sans-serif; font-weight: 700; font-size: 13px; margin: 18px 0 8px; text-transform: uppercase; letter-spacing: 0.08em; color: #475569; }
  .kpi-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 20px; }
  .kpi-grid-4 { grid-template-columns: repeat(4, 1fr); }
  .kpi { padding: 18px; border-radius: 10px; border: 1px solid #e2e8f0; }
  .kpi-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.08em; color: #64748b; font-weight: 700; margin-bottom: 6px; }
  .kpi-value { font-family: 'Manrope', sans-serif; font-size: 26px; font-weight: 800; line-height: 1.1; letter-spacing: -0.02em; color: #0f172a; }
  .kpi-sub { font-size: 11px; color: #64748b; margin-top: 4px; }
  .kpi-emerald { background: linear-gradient(135deg, #ecfdf5, #d1fae5); border-color: #6ee7b7; }
  .kpi-emerald .kpi-value { color: #0F172A; }
  .kpi-rose { background: linear-gradient(135deg, #fff1f2, #ffe4e6); border-color: #fda4af; }
  .kpi-rose .kpi-value { color: #be123c; }
  .kpi-amber { background: linear-gradient(135deg, #fffbeb, #fef3c7); border-color: #fcd34d; }
  .kpi-amber .kpi-value { color: #92400e; }
  .kpi-sky { background: linear-gradient(135deg, #f0f9ff, #e0f2fe); border-color: #7dd3fc; }
  .kpi-sky .kpi-value { color: #075985; }
  .kpi-violet { background: linear-gradient(135deg, #f5f3ff, #ede9fe); border-color: #c4b5fd; }
  .kpi-violet .kpi-value { color: #0f172a; }
  .kpi-slate { background: #0f172a; border-color: #0f172a; }
  .kpi-slate .kpi-label { color: #94a3b8; }
  .kpi-slate .kpi-value { color: white; }
  .kpi-slate .kpi-sub { color: #94a3b8; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 16px; font-size: 12px; }
  th { text-align: left; padding: 10px 12px; background: #f8fafc; font-size: 10px; text-transform: uppercase; letter-spacing: 0.08em; color: #475569; border-bottom: 2px solid #cbd5e1; font-weight: 700; }
  th.right, td.right { text-align: right; }
  th.center, td.center { text-align: center; }
  td { padding: 10px 12px; border-bottom: 1px solid #f1f5f9; }
  tr:nth-child(even) td { background: #fafbfc; }
  .total-row td { border-top: 2px solid #0f172a; border-bottom: 2px solid #0f172a; font-weight: 700; background: #f1f5f9 !important; }
  td.mono, .mono { font-family: 'JetBrains Mono', monospace; font-variant-numeric: tabular-nums; }
  .green, .profit { color: #1D4ED8; font-weight: 700; }
  .red, .loss { color: #e11d48; font-weight: 700; }
  .amber, .expense { color: #475569; font-weight: 700; }
  .violet, .vat { color: #7c3aed; font-weight: 700; }
  .sky, .balance { color: #0284c7; font-weight: 700; }
  .badge { display: inline-block; padding: 3px 10px; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; border-radius: 4px; border: 1px solid; }
  .badge-safe { background: #ecfdf5; color: #0F172A; border-color: #6ee7b7; }
  .badge-caution { background: #fffbeb; color: #92400e; border-color: #fcd34d; }
  .badge-risky { background: #fff1f2; color: #be123c; border-color: #fda4af; }
  .badge-neutral { background: #f1f5f9; color: #475569; border-color: #cbd5e1; }
  .risk-banner { padding: 16px; border-radius: 10px; border: 1px solid; margin-bottom: 24px; }
  .risk-safe { background: #ecfdf5; border-color: #6ee7b7; }
  .risk-caution { background: #fffbeb; border-color: #fcd34d; }
  .risk-risky { background: #fff1f2; border-color: #fda4af; }
  .risk-title { font-family: 'Manrope', sans-serif; font-weight: 700; font-size: 14px; margin-bottom: 8px; }
  .risk-list { font-size: 12px; padding-left: 18px; }
  .risk-list li { margin-bottom: 3px; }
  .footer { margin-top: 60px; padding-top: 16px; border-top: 1px solid #e2e8f0; font-size: 10px; color: #94a3b8; text-align: center; letter-spacing: 0.05em; }
  .footer strong { color: #475569; }
  .print-btn { position: fixed; top: 16px; right: 16px; padding: 10px 18px; background: #1D4ED8; color: white; border: none; border-radius: 8px; font-weight: 700; cursor: pointer; font-size: 13px; box-shadow: 0 4px 12px rgba(5, 150, 105, 0.25); }
  .print-btn:hover { background: #15803D; }
  .progress-bar { height: 8px; background: #e2e8f0; border-radius: 4px; overflow: hidden; margin: 6px 0; }
  .progress-fill { height: 100%; background: linear-gradient(90deg, #1D4ED8, #2563EB); }
  @media print {
    body { padding: 20px; }
    .no-print { display: none !important; }
    .kpi-slate { background: #0f172a !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .kpi-emerald, .kpi-rose, .kpi-amber, .kpi-sky, .kpi-violet { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    h2 { page-break-after: avoid; }
    table { page-break-inside: avoid; }
  }
`;

const reportMoney = (n, currency = 'KES', short = false) => {
  if (n === null || n === undefined || isNaN(n)) return '—';
  if (short) {
    const abs = Math.abs(n);
    const sign = n < 0 ? '-' : '';
    if (abs >= 1_000_000) return `${sign}${currency} ${(abs / 1_000_000).toFixed(2)}M`;
    if (abs >= 1_000) return `${sign}${currency} ${(abs / 1_000).toFixed(1)}K`;
  }
  return `${currency} ${Number(n).toLocaleString('en-US', { maximumFractionDigits: 0 })}`;
};

const buildProjectReportHTML = (project) => {
  const c = calcProject(project);
  const ds = computeDocStatus(project);
  const hs = computeHealthScore(project, c);
  const cur = project.currency || 'KES';

  const docRows = (project.documents || []).map(d => {
    const t = d.templateId ? getDocTemplate(d.templateId) : null;
    const label = t?.label || d.customLabel || 'Document';
    const statusBadge = `<span class="badge badge-${d.status === 'received' ? 'safe' : d.status === 'expired' ? 'risky' : 'neutral'}">${(d.status || '').toUpperCase()}</span>`;
    return `<tr><td>${label}</td><td>${t?.category || d.category || '—'}</td><td class="center">${statusBadge}</td><td>${d.expiryDate || '—'}</td><td>${d.fileName ? '✓ Uploaded' : d.url ? '✓ Linked' : '—'}</td></tr>`;
  }).join('');

  const partnerRows = c.partnersWithShare.map(pt => `
    <tr>
      <td>${pt.name || 'Unnamed'}</td>
      <td class="right mono">${fmtPct(pt.sharePct)}</td>
      <td class="right mono">${reportMoney(pt.payout, cur)}</td>
      <td class="right mono ${pt.profitShare >= 0 ? 'green' : 'red'}">${reportMoney(pt.profitShare, cur)}</td>
    </tr>
  `).join('');

  const deductionRows = [
    ['Mobilisation', +project.mobilisationCost || 0],
    ['Performance Bond', +project.performanceBondCost || 0],
    ['Bank Guarantee', +project.bankGuaranteeCost || 0],
    ['Insurance', +project.insuranceCost || 0],
    ['Tender / Application', +project.tenderCost || 0],
    ['Legal / Documentation', +project.legalCost || 0],
    ['Commissions', +project.commissionCost || 0],
    ...(project.customDeductions || []).map(d => [d.label || 'Custom', +d.amount || 0]),
    ...(project.partners || []).filter(pt => +pt.payout > 0).map(pt => [`Payout — ${pt.name || 'Partner'}`, +pt.payout || 0])
  ].filter(([, v]) => v > 0);

  return `
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>${project.name || 'Project Report'}</title><style>${REPORT_CSS}</style></head>
<body>
<button class="print-btn no-print" onclick="window.print()">📄 Save as PDF</button>

<div class="brand-bar">
  <div class="brand-name">RAYS INVESTCALC <span style="color:#1D4ED8">PRO</span></div>
  <div class="brand-tag">Project Investment Report</div>
</div>

<div class="doc-title">${project.name || 'Untitled Project'}</div>
<div class="doc-subtitle">${project.client || 'No client specified'}</div>
<div class="doc-meta">
  ${project.location ? `<div><strong>Location</strong>${project.location}</div>` : ''}
  ${project.type ? `<div><strong>Type</strong>${project.type}</div>` : ''}
  <div><strong>Status</strong><span class="badge badge-neutral">${project.status?.toUpperCase()}</span></div>
  <div><strong>Risk</strong><span class="badge badge-${c.riskLevel}">${c.riskLevel?.toUpperCase()}</span></div>
  <div><strong>Health Score</strong>${hs.score} / 100</div>
  <div><strong>Generated</strong>${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}</div>
</div>

<h2>Executive Summary</h2>
<div class="kpi-grid kpi-grid-4">
  <div class="kpi kpi-slate">
    <div class="kpi-label">Contract Value</div>
    <div class="kpi-value">${reportMoney(c.contract, cur, true)}</div>
    <div class="kpi-sub">Gross before VAT</div>
  </div>
  <div class="kpi kpi-sky">
    <div class="kpi-label">Real Cash Upfront</div>
    <div class="kpi-value">${reportMoney(c.realCashAfterDeductions, cur, true)}</div>
    <div class="kpi-sub">After deductions</div>
  </div>
  <div class="kpi kpi-${c.projectedFinalProfit >= 0 ? 'emerald' : 'rose'}">
    <div class="kpi-label">Final Profit</div>
    <div class="kpi-value">${reportMoney(c.projectedFinalProfit, cur, true)}</div>
    <div class="kpi-sub">Margin: ${c.profitMargin.toFixed(2)}%</div>
  </div>
  <div class="kpi kpi-amber">
    <div class="kpi-label">Total Deductions</div>
    <div class="kpi-value">${reportMoney(c.totalDeductions, cur, true)}</div>
    <div class="kpi-sub">${((c.totalDeductions / c.contract) * 100).toFixed(1)}% of contract</div>
  </div>
</div>

<h2>Risk Assessment</h2>
<div class="risk-banner risk-${c.riskLevel}">
  <div class="risk-title">${c.riskLevel === 'safe' ? '✓ Project appears financially sound' : c.riskLevel === 'caution' ? '⚠ Proceed with caution' : '✕ High risk — review carefully'}</div>
  <ul class="risk-list">${c.riskReasons.map(r => `<li>${r}</li>`).join('')}</ul>
</div>

<h2>Financial Walkthrough</h2>
<table>
  <tbody>
    <tr><td>Contract value</td><td class="right mono">${reportMoney(c.contract, cur)}</td></tr>
    <tr><td>VAT (${fmtPct(project.vatPct)})</td><td class="right mono violet">− ${reportMoney(c.vatAmount, cur)}</td></tr>
    <tr><td><strong>Net after VAT</strong></td><td class="right mono"><strong>${reportMoney(c.netAfterVat, cur)}</strong></td></tr>
    <tr><td>Expected profit (${fmtPct(project.profitPct)})</td><td class="right mono amber">${reportMoney(c.expectedProfit, cur)}</td></tr>
    <tr><td>Advance (${fmtPct(project.advancePct)})</td><td class="right mono">${reportMoney(c.advanceAmount, cur)}</td></tr>
    <tr><td>Bank retention (${fmtPct(project.bankRetentionPct)})</td><td class="right mono red">− ${reportMoney(c.bankRetained, cur)}</td></tr>
    <tr><td><strong>Advance cash available</strong></td><td class="right mono sky"><strong>${reportMoney(c.advanceCashAvailable, cur)}</strong></td></tr>
    <tr><td>Total deductions</td><td class="right mono red">− ${reportMoney(c.totalDeductions, cur)}</td></tr>
    <tr class="total-row"><td>Real cash after deductions</td><td class="right mono ${c.realCashAfterDeductions >= 0 ? 'green' : 'red'}">${reportMoney(c.realCashAfterDeductions, cur)}</td></tr>
    <tr class="total-row"><td>Projected final profit</td><td class="right mono ${c.projectedFinalProfit >= 0 ? 'green' : 'red'}">${reportMoney(c.projectedFinalProfit, cur)}</td></tr>
  </tbody>
</table>

${deductionRows.length > 0 ? `
<h2>Deductions Breakdown</h2>
<table>
  <thead><tr><th>Item</th><th class="right">Amount (${cur})</th></tr></thead>
  <tbody>
    ${deductionRows.map(([label, v]) => `<tr><td>${label}</td><td class="right mono">${reportMoney(v, cur)}</td></tr>`).join('')}
    <tr class="total-row"><td>Total Deductions</td><td class="right mono red">${reportMoney(c.totalDeductions, cur)}</td></tr>
  </tbody>
</table>` : ''}

${partnerRows ? `
<h2>Partner Profit Share</h2>
<table>
  <thead><tr><th>Partner</th><th class="right">Share</th><th class="right">Direct Payout</th><th class="right">Profit Share</th></tr></thead>
  <tbody>${partnerRows}</tbody>
</table>` : ''}

${docRows ? `
<h2>Documents Status</h2>
<div style="margin-bottom:12px;font-size:12px;">
  <strong>Compliance:</strong> ${ds.criticalReceived.length}/${ds.criticalTemplates.length} critical (${ds.compliancePct.toFixed(0)}%)
  <div class="progress-bar"><div class="progress-fill" style="width:${ds.compliancePct}%"></div></div>
</div>
<table>
  <thead><tr><th>Document</th><th>Category</th><th class="center">Status</th><th>Expiry</th><th>Attached</th></tr></thead>
  <tbody>${docRows}</tbody>
</table>` : ''}

${project.notes ? `<h2>Notes</h2><div style="font-size:13px;line-height:1.6;background:#fffbeb;border-left:4px solid #fcd34d;padding:14px 16px;border-radius:4px;white-space:pre-wrap;">${project.notes}</div>` : ''}

<div class="footer">
  Generated by <strong>Rays InvestCalc Pro</strong> · ${new Date().toLocaleString()} · Project ID: ${project.id}
</div>
</body></html>`;
};

const buildPortfolioReportHTML = (projects, title = 'Portfolio Report') => {
  const calcs = projects.map(p => ({ p, c: calcProject(p) }));
  const totalContract = calcs.reduce((s, x) => s + x.c.contract, 0);
  const totalProfit = calcs.reduce((s, x) => s + x.c.projectedFinalProfit, 0);
  const totalCash = calcs.reduce((s, x) => s + x.c.realCashAfterDeductions, 0);
  const totalDeductions = calcs.reduce((s, x) => s + x.c.totalDeductions, 0);
  const safe = calcs.filter(x => x.c.riskLevel === 'safe').length;
  const caution = calcs.filter(x => x.c.riskLevel === 'caution').length;
  const risky = calcs.filter(x => x.c.riskLevel === 'risky').length;
  const recs = computeRecommendations(calcs);

  const projectRows = calcs.map(({ p, c }) => `
    <tr>
      <td><strong>${p.name || 'Untitled'}</strong><br><span style="color:#64748b;font-size:11px;">${p.client || '—'}</span></td>
      <td><span class="badge badge-neutral">${p.status?.toUpperCase()}</span></td>
      <td class="right mono">${reportMoney(c.contract, p.currency, true)}</td>
      <td class="right mono ${c.realCashAfterDeductions >= 0 ? 'green' : 'red'}">${reportMoney(c.realCashAfterDeductions, p.currency, true)}</td>
      <td class="right mono ${c.projectedFinalProfit >= 0 ? 'green' : 'red'}">${reportMoney(c.projectedFinalProfit, p.currency, true)}</td>
      <td class="right mono">${c.profitMargin.toFixed(1)}%</td>
      <td class="center"><span class="badge badge-${c.riskLevel}">${c.riskLevel?.toUpperCase()}</span></td>
    </tr>
  `).join('');

  return `
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>${title}</title><style>${REPORT_CSS}</style></head>
<body>
<button class="print-btn no-print" onclick="window.print()">📄 Save as PDF</button>

<div class="brand-bar">
  <div class="brand-name">RAYS INVESTCALC <span style="color:#1D4ED8">PRO</span></div>
  <div class="brand-tag">${title}</div>
</div>

<div class="doc-title">${title}</div>
<div class="doc-subtitle">${projects.length} project${projects.length === 1 ? '' : 's'} analyzed</div>
<div class="doc-meta">
  <div><strong>Generated</strong>${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}</div>
  <div><strong>Reporting Period</strong>All-time</div>
</div>

<h2>Portfolio KPIs</h2>
<div class="kpi-grid kpi-grid-4">
  <div class="kpi kpi-slate">
    <div class="kpi-label">Projects</div>
    <div class="kpi-value">${projects.length}</div>
    <div class="kpi-sub">In portfolio</div>
  </div>
  <div class="kpi kpi-emerald">
    <div class="kpi-label">Combined Contracts</div>
    <div class="kpi-value">${reportMoney(totalContract, 'KES', true)}</div>
    <div class="kpi-sub">Total value</div>
  </div>
  <div class="kpi kpi-${totalProfit >= 0 ? 'emerald' : 'rose'}">
    <div class="kpi-label">Projected Profit</div>
    <div class="kpi-value">${reportMoney(totalProfit, 'KES', true)}</div>
    <div class="kpi-sub">After all deductions</div>
  </div>
  <div class="kpi kpi-amber">
    <div class="kpi-label">Total Deductions</div>
    <div class="kpi-value">${reportMoney(totalDeductions, 'KES', true)}</div>
    <div class="kpi-sub">Costs & retentions</div>
  </div>
</div>

<h2>Risk Distribution</h2>
<div class="kpi-grid">
  <div class="kpi kpi-emerald"><div class="kpi-label">Safe</div><div class="kpi-value">${safe}</div></div>
  <div class="kpi kpi-amber"><div class="kpi-label">Caution</div><div class="kpi-value">${caution}</div></div>
  <div class="kpi kpi-rose"><div class="kpi-label">Risky</div><div class="kpi-value">${risky}</div></div>
  <div class="kpi kpi-sky"><div class="kpi-label">Real Cash Total</div><div class="kpi-value">${reportMoney(totalCash, 'KES', true)}</div></div>
</div>

<h2>Project Ledger</h2>
<table>
  <thead><tr><th>Project</th><th>Status</th><th class="right">Contract</th><th class="right">Real Cash</th><th class="right">Final Profit</th><th class="right">Margin</th><th class="center">Risk</th></tr></thead>
  <tbody>${projectRows}</tbody>
</table>

${recs.length > 0 ? `
<h2>Recommendations</h2>
<table>
  <thead><tr><th>Priority</th><th>Recommendation</th><th>Detail</th></tr></thead>
  <tbody>${recs.slice(0, 10).map(r => `
    <tr>
      <td class="center"><span class="badge badge-${r.priority === 'high' ? 'risky' : 'caution'}">${r.priority.toUpperCase()}</span></td>
      <td><strong>${r.title}</strong></td>
      <td style="color:#475569;font-size:11px;">${r.body}</td>
    </tr>
  `).join('')}</tbody>
</table>` : ''}

<div class="footer">
  Generated by <strong>Rays InvestCalc Pro</strong> · ${new Date().toLocaleString()} · ${projects.length} projects analyzed
</div>
</body></html>`;
};

const openReportPDF = (html) => {
  const w = window.open('', '_blank', 'width=900,height=1000');
  if (!w) {
    alert('Please allow pop-ups for this site so reports can open.');
    return;
  }
  w.document.write(html);
  w.document.close();
};

// ====== EXCEL GENERATORS ======

const buildProjectXLSX = (project) => {
  const c = calcProject(project);
  const ds = computeDocStatus(project);
  const hs = computeHealthScore(project, c);
  const cur = project.currency || 'KES';
  const wb = XLSX.utils.book_new();

  // Summary sheet
  const summary = [
    ['RAYS INVESTCALC PRO — PROJECT REPORT'],
    [],
    ['Project Name', project.name || ''],
    ['Client', project.client || ''],
    ['Location', project.location || ''],
    ['Type', project.type || ''],
    ['Status', project.status || ''],
    ['Currency', cur],
    ['Generated', new Date().toLocaleString()],
    [],
    ['EXECUTIVE SUMMARY'],
    ['Metric', `Amount (${cur})`],
    ['Contract Value', c.contract],
    ['VAT Amount', c.vatAmount],
    ['Net after VAT', c.netAfterVat],
    ['Expected Profit', c.expectedProfit],
    ['Advance Amount', c.advanceAmount],
    ['Bank Retention', c.bankRetained],
    ['Advance Cash Available', c.advanceCashAvailable],
    ['Total Deductions', c.totalDeductions],
    ['Real Cash After Deductions', c.realCashAfterDeductions],
    ['Final Projected Profit', c.projectedFinalProfit],
    [],
    ['Effective Margin %', c.profitMargin],
    ['Capital Efficiency (x)', c.capitalEfficiency],
    ['Risk Level', c.riskLevel],
    ['Health Score', hs.score]
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(summary);
  ws1['!cols'] = [{ wch: 30 }, { wch: 22 }];
  XLSX.utils.book_append_sheet(wb, ws1, 'Summary');

  // Deductions sheet
  const deductions = [
    ['DEDUCTIONS BREAKDOWN'],
    [],
    ['Item', `Amount (${cur})`],
    ['Mobilisation', +project.mobilisationCost || 0],
    ['Performance Bond', +project.performanceBondCost || 0],
    ['Bank Guarantee', +project.bankGuaranteeCost || 0],
    ['Insurance', +project.insuranceCost || 0],
    ['Tender / Application', +project.tenderCost || 0],
    ['Legal / Documentation', +project.legalCost || 0],
    ['Commissions', +project.commissionCost || 0],
    ...(project.customDeductions || []).map(d => [d.label || 'Custom', +d.amount || 0]),
    ...(project.partners || []).filter(pt => +pt.payout > 0).map(pt => [`Partner Payout — ${pt.name || 'Partner'}`, +pt.payout || 0]),
    [],
    ['TOTAL', c.totalDeductions]
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(deductions);
  ws2['!cols'] = [{ wch: 35 }, { wch: 22 }];
  XLSX.utils.book_append_sheet(wb, ws2, 'Deductions');

  // Partners sheet
  if (c.partnersWithShare.length > 0) {
    const partners = [
      ['PARTNER PROFIT SHARES'],
      [],
      ['Partner', 'Share %', `Direct Payout (${cur})`, `Profit Share (${cur})`, `Total to Partner (${cur})`],
      ...c.partnersWithShare.map(pt => [pt.name || 'Unnamed', pt.sharePct, pt.payout, pt.profitShare, +pt.payout + pt.profitShare])
    ];
    const ws3 = XLSX.utils.aoa_to_sheet(partners);
    ws3['!cols'] = [{ wch: 25 }, { wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, ws3, 'Partners');
  }

  // Documents sheet
  if ((project.documents || []).length > 0) {
    const docs = [
      ['DOCUMENTS REGISTER'],
      [`Compliance: ${ds.criticalReceived.length}/${ds.criticalTemplates.length} critical (${ds.compliancePct.toFixed(0)}%)`],
      [],
      ['Document', 'Category', 'Status', 'Issue Date', 'Expiry Date', 'Link', 'File', 'Notes'],
      ...(project.documents || []).map(d => {
        const t = d.templateId ? getDocTemplate(d.templateId) : null;
        return [
          t?.label || d.customLabel || 'Document',
          t?.category || d.category || '—',
          d.status || 'pending',
          d.issuedDate || '',
          d.expiryDate || '',
          d.url || '',
          d.fileName || '',
          d.notes || ''
        ];
      })
    ];
    const ws4 = XLSX.utils.aoa_to_sheet(docs);
    ws4['!cols'] = [{ wch: 32 }, { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 30 }, { wch: 24 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, ws4, 'Documents');
  }

  XLSX.writeFile(wb, `${(project.name || 'project').replace(/[^a-z0-9]/gi, '_')}_report_${new Date().toISOString().split('T')[0]}.xlsx`);
};

const buildPortfolioXLSX = (projects, title = 'Portfolio Report') => {
  const calcs = projects.map(p => ({ p, c: calcProject(p) }));
  const wb = XLSX.utils.book_new();

  // Summary sheet
  const summary = [
    [`RAYS INVESTCALC PRO — ${title.toUpperCase()}`],
    ['Generated', new Date().toLocaleString()],
    ['Projects', projects.length],
    [],
    ['Project', 'Client', 'Status', 'Currency', 'Contract', 'VAT', 'Expected Profit', 'Total Deductions', 'Advance Cash', 'Real Cash', 'Final Profit', 'Margin %', 'Risk', 'Health'],
    ...calcs.map(({ p, c }) => [
      p.name || 'Untitled', p.client || '', p.status || '', p.currency || 'KES',
      c.contract, c.vatAmount, c.expectedProfit, c.totalDeductions, c.advanceCashAvailable, c.realCashAfterDeductions, c.projectedFinalProfit, c.profitMargin.toFixed(2), c.riskLevel, computeHealthScore(p, c).score
    ])
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(summary);
  ws1['!cols'] = [{ wch: 30 }, { wch: 22 }, { wch: 12 }, { wch: 10 }, ...Array(10).fill({ wch: 16 })];
  XLSX.utils.book_append_sheet(wb, ws1, 'Summary');

  // Cashflow sheet
  const cashflow = [
    ['CASHFLOW BREAKDOWN'],
    [],
    ['Project', 'Currency', 'Contract', 'Advance', 'Bank Retention', 'Advance Cash', 'Total Deductions', 'Real Cash', 'Final Profit'],
    ...calcs.map(({ p, c }) => [
      p.name || 'Untitled', p.currency || 'KES',
      c.contract, c.advanceAmount, c.bankRetained, c.advanceCashAvailable, c.totalDeductions, c.realCashAfterDeductions, c.projectedFinalProfit
    ])
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(cashflow);
  ws2['!cols'] = [{ wch: 30 }, { wch: 10 }, ...Array(7).fill({ wch: 16 })];
  XLSX.utils.book_append_sheet(wb, ws2, 'Cashflow');

  // Deductions sheet
  const ded = [
    ['DEDUCTIONS PER PROJECT'],
    [],
    ['Project', 'Mobilisation', 'Performance Bond', 'Bank Guarantee', 'Insurance', 'Tender', 'Legal', 'Commissions', 'Custom Total', 'Partner Payouts', 'Total'],
    ...calcs.map(({ p, c }) => [
      p.name || 'Untitled',
      +p.mobilisationCost || 0, +p.performanceBondCost || 0, +p.bankGuaranteeCost || 0, +p.insuranceCost || 0,
      +p.tenderCost || 0, +p.legalCost || 0, +p.commissionCost || 0, c.customTotal, c.partnerTotal, c.totalDeductions
    ])
  ];
  const ws3 = XLSX.utils.aoa_to_sheet(ded);
  ws3['!cols'] = [{ wch: 30 }, ...Array(10).fill({ wch: 16 })];
  XLSX.utils.book_append_sheet(wb, ws3, 'Deductions');

  // Partners sheet
  const partnerRows = [];
  calcs.forEach(({ p, c }) => {
    c.partnersWithShare.forEach(pt => {
      partnerRows.push([p.name || 'Untitled', pt.name || 'Unnamed', pt.sharePct, pt.payout, pt.profitShare, +pt.payout + pt.profitShare]);
    });
  });
  if (partnerRows.length > 0) {
    const partners = [
      ['PARTNER SHARES ACROSS PORTFOLIO'],
      [],
      ['Project', 'Partner', 'Share %', 'Direct Payout', 'Profit Share', 'Total to Partner'],
      ...partnerRows
    ];
    const ws4 = XLSX.utils.aoa_to_sheet(partners);
    ws4['!cols'] = [{ wch: 30 }, { wch: 25 }, { wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, ws4, 'Partners');
  }

  // Documents sheet
  const docRows = [];
  calcs.forEach(({ p }) => {
    (p.documents || []).forEach(d => {
      const t = d.templateId ? getDocTemplate(d.templateId) : null;
      docRows.push([
        p.name || 'Untitled',
        t?.label || d.customLabel || 'Document',
        t?.category || d.category || '—',
        d.status || 'pending',
        d.issuedDate || '',
        d.expiryDate || '',
        d.url || '',
        d.fileName || '',
        d.notes || ''
      ]);
    });
  });
  if (docRows.length > 0) {
    const docs = [
      ['DOCUMENTS REGISTER ACROSS PORTFOLIO'],
      [],
      ['Project', 'Document', 'Category', 'Status', 'Issue Date', 'Expiry', 'Link', 'File', 'Notes'],
      ...docRows
    ];
    const ws5 = XLSX.utils.aoa_to_sheet(docs);
    ws5['!cols'] = [{ wch: 30 }, { wch: 32 }, { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 30 }, { wch: 24 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, ws5, 'Documents');
  }

  XLSX.writeFile(wb, `rays_${title.toLowerCase().replace(/[^a-z0-9]/gi, '_')}_${new Date().toISOString().split('T')[0]}.xlsx`);
};





// Centralized definitions of every column / metric in the app

// ============================================================
// UI PRIMITIVES
// ============================================================
const Card = ({ children, className = '', padding = 'p-6' }) => (
  <div className={`${padding} ${className}`} style={{
    background: '#FFFFFF',
    border: '1px solid #D8E4F8',
    borderRadius: '20px',
    boxShadow: '0 10px 28px rgba(37, 99, 235, 0.08)'
  }}>{children}</div>
);

const KpiTile = ({ label, value, sub, icon: Icon, tone = 'slate', onClick, help, trend }) => {
  // Premium spec: blue circle icons, red for risky only
  const isRisky = tone === 'rose';
  const iconBg = isRisky ? '#FFE4E6' : '#DBEAFE';
  const iconColor = isRisky ? '#F43F5E' : '#2563EB';
  return (
    <button
      onClick={onClick}
      disabled={!onClick}
      className={`group relative bg-white text-left w-full transition-all ${onClick ? 'cursor-pointer hover:-translate-y-0.5' : 'cursor-default'}`}
      style={{
        border: '1px solid #D8E4F8',
        borderRadius: '20px',
        padding: '22px',
        boxShadow: '0 10px 28px rgba(37, 99, 235, 0.08)'
      }}
      onMouseEnter={e => onClick && (e.currentTarget.style.boxShadow = '0 16px 36px rgba(37, 99, 235, 0.14)')}
      onMouseLeave={e => onClick && (e.currentTarget.style.boxShadow = '0 10px 28px rgba(37, 99, 235, 0.08)')}
    >
      <div className="flex items-start justify-between mb-4">
        <div className="rounded-full flex items-center justify-center flex-shrink-0" style={{ background: iconBg, width: '40px', height: '40px' }}>
          <Icon size={18} strokeWidth={2.4} style={{ color: iconColor }} />
        </div>
        {help && (
          <Tooltip text={help} position="bottom">
            <span className="p-1 rounded-md cursor-help inline-flex" style={{ color: '#94A3B8' }}>
              <HelpCircle size={12} strokeWidth={2.4} />
            </span>
          </Tooltip>
        )}
      </div>
      <div className="text-[11px] uppercase tracking-[0.10em] font-bold mb-1.5" style={{ color: '#0F172A' }}>{label}</div>
      <div className="num-display text-[24px] leading-none mb-2" style={{ color: '#071739' }}>{value}</div>
      <div className="flex items-center justify-between mt-2">
        {sub && <div className="text-[12px] font-semibold" style={{ color: '#475569' }}>{sub}</div>}
        {trend && (
          <div className="text-[12px] font-extrabold flex items-center gap-1" style={{ color: trend.positive ? '#16A34A' : '#F43F5E' }}>
            {trend.positive ? <TrendingUp size={12} strokeWidth={2.8} /> : <TrendingDown size={12} strokeWidth={2.8} />}
            {trend.label}
          </div>
        )}
        {onClick && !trend && <ChevronRight size={14} className="group-hover:translate-x-0.5 transition-all ml-auto" strokeWidth={2.4} style={{ color: '#CBD5E1' }} />}
      </div>
    </button>
  );
};

const Pill = ({ children, tone = 'neutral', size = 'md' }) => {
  const tones = {
    safe: 'bg-[#EFF6FF] text-[#2563EB] border-[#BFDBFE]',
    caution: 'bg-slate-100 text-slate-700 border-slate-300',
    risky: 'bg-red-50 text-red-700 border-red-200',
    neutral: 'bg-slate-100 text-slate-700 border-slate-200',
    planning: 'bg-[#EFF6FF] text-[#2563EB] border-[#BFDBFE]',
    active: 'bg-[#0F172A] text-white border-[#0F172A]',
    completed: 'bg-slate-100 text-slate-600 border-slate-200',
    cancelled: 'bg-red-50 text-red-600 border-red-200'
  };
  const sizes = { sm: 'px-2 py-0.5 text-[10px]', md: 'px-2.5 py-1 text-[11px]' };
  return (
    <span className={`inline-flex items-center gap-1.5 border rounded-md font-bold uppercase tracking-wider ${tones[tone]} ${sizes[size]}`}>
      {children}
    </span>
  );
};

const Input = ({ label, value, onChange, type = 'text', prefix, suffix, hint, required, placeholder, help }) => (
  <div>
    <label className="block text-xs font-bold text-slate-900 mb-1.5">
      {label} {required && <span className="text-rose-500">*</span>}
      {help && <InfoIcon text={help} />}
    </label>
    <div className="flex items-stretch rounded-lg bg-white border border-[#CBD5E1] focus-within:border-blue-500 focus-within:ring-2 focus-within:ring-blue-200 transition-all">
      {prefix && <div className="bg-[#EFF6FF] border-r border-[#D8E4F8] px-3 flex items-center text-sm text-[#2563EB] font-semibold rounded-l-lg">{prefix}</div>}
      <input
        type={type}
        value={value ?? ''}
        placeholder={placeholder}
        onChange={e => onChange(type === 'number' ? (e.target.value === '' ? '' : +e.target.value) : e.target.value)}
        className="flex-1 min-w-0 bg-transparent px-3 py-2.5 text-sm text-slate-900 placeholder:text-blue-700 focus:outline-none rounded-lg"
      />
      {suffix && <div className="bg-[#EFF6FF] border-l border-[#D8E4F8] px-3 flex items-center text-sm text-[#2563EB] font-semibold rounded-r-lg">{suffix}</div>}
    </div>
    {hint && <div className="text-xs text-blue-700 mt-1">{hint}</div>}
  </div>
);

// PctAmountInput — toggle between entering a percentage OR a fixed amount
// storeAs='pct'    -> value is a percentage; amount is computed from base
// storeAs='amount' -> value is an amount; percentage is computed from base
const PctAmountInput = ({ label, value, onChange, base, currency = 'KES', hint, help, required, baseLabel = 'base', storeAs = 'pct' }) => {
  const [mode, setMode] = useState(storeAs);
  const v = +value || 0;
  const baseNum = +base || 0;

  const pct = storeAs === 'pct' ? v : (baseNum > 0 ? (v / baseNum) * 100 : 0);
  const amount = storeAs === 'amount' ? v : (baseNum > 0 ? baseNum * (v / 100) : 0);

  // Compute the display string for a given mode
  const computeDisplay = (m) => {
    const pctNow = storeAs === 'pct' ? v : (baseNum > 0 ? (v / baseNum) * 100 : 0);
    const amtNow = storeAs === 'amount' ? v : (baseNum > 0 ? baseNum * (v / 100) : 0);
    if (m === 'pct') return pctNow === 0 ? '' : String(Math.round(pctNow * 10000) / 10000);
    return amtNow === 0 ? '' : String(Math.round(amtNow));
  };

  // Local string state — what the user actually sees while typing
  const [inputStr, setInputStr] = useState(() => computeDisplay(storeAs));

  // When mode changes, refresh the display from current props
  const prevModeRef = useRef(mode);
  useEffect(() => {
    if (prevModeRef.current !== mode) {
      setInputStr(computeDisplay(mode));
      prevModeRef.current = mode;
    }
  }, [mode, v, baseNum]);

  const handleChange = (raw) => {
    setInputStr(raw); // preserve exactly what the user typed (decimals etc.)
    if (raw === '' || raw === '-' || raw === '.') {
      onChange(0);
      return;
    }
    const num = Number(raw);
    if (Number.isNaN(num)) return;
    if (mode === storeAs) {
      onChange(num);
    } else if (mode === 'pct') {
      // user entered %, but we store amount: convert
      if (baseNum > 0) onChange(baseNum * num / 100);
    } else {
      // user entered amount, but we store %: convert
      if (baseNum > 0) onChange((num / baseNum) * 100);
    }
  };

  const switchTo = (newMode) => {
    if (newMode === mode) return;
    if (newMode === 'amount' && baseNum === 0) return;
    setMode(newMode);
  };

  const canSwitchAmount = baseNum > 0;

  return (
    <div>
      <div className="flex items-center justify-between mb-1.5 gap-2 flex-wrap">
        <label className="text-xs font-bold text-slate-900 flex items-center">
          {label} {required && <span className="text-rose-500">*</span>}
          {help && <InfoIcon text={help} />}
        </label>
        {/* Larger, clearer segmented control */}
        <div className="inline-flex border border-slate-300 rounded-lg overflow-hidden text-xs font-bold shadow-sm">
          <button
            type="button"
            onClick={() => switchTo('pct')}
            className={`px-4 py-1.5 transition-colors ${mode === 'pct' ? 'bg-[#EFF6FF]0 text-slate-900' : 'bg-transparent text-slate-900 hover:bg-[#F1F5F9] active:bg-white/[0.1]'}`}
          >%</button>
          <button
            type="button"
            onClick={() => switchTo('amount')}
            disabled={!canSwitchAmount}
            title={!canSwitchAmount ? 'Set the base value first' : `Switch to ${currency}`}
            className={`px-4 py-1.5 transition-colors border-l border-slate-300 ${mode === 'amount' ? 'bg-[#EFF6FF]0 text-slate-900' : canSwitchAmount ? 'bg-transparent text-slate-900 hover:bg-[#F1F5F9] active:bg-white/[0.1]' : 'bg-[#EFF6FF] text-blue-600 cursor-not-allowed'}`}
          >{currency}</button>
        </div>
      </div>
      <div className="flex items-stretch rounded-lg bg-white border border-[#CBD5E1] focus-within:border-blue-500 focus-within:ring-2 focus-within:ring-blue-200 transition-all">
        {mode === 'amount' && <div className="bg-[#EFF6FF] border-r border-[#D8E4F8] px-3 flex items-center text-sm text-[#2563EB] font-semibold rounded-l-lg">{currency}</div>}
        <input
          type="number"
          inputMode="decimal"
          value={inputStr}
          step="any"
          onChange={e => handleChange(e.target.value)}
          className="flex-1 min-w-0 bg-transparent px-3 py-2.5 text-sm text-slate-900 placeholder:text-blue-700 focus:outline-none rounded-lg"
        />
        {mode === 'pct' && <div className="bg-[#EFF6FF] border-l border-[#D8E4F8] px-3 flex items-center text-sm text-[#2563EB] font-semibold rounded-r-lg">%</div>}
      </div>
      <div className="flex items-center justify-between mt-1 gap-2 flex-wrap">
        {hint && <div className="text-xs text-[#2563EB]">{hint}</div>}
        {baseNum > 0 && (
          <div className="text-[11px] text-blue-700 font-mono tabular-nums ml-auto">
            {mode === 'pct'
              ? `≈ ${currency} ${amount.toLocaleString('en-US', { maximumFractionDigits: 0 })} ${baseLabel}`
              : `≈ ${pct.toFixed(2)}% ${baseLabel}`}
          </div>
        )}
      </div>
    </div>
  );
};

const Select = ({ label, value, onChange, options }) => (
  <div>
    <label className="block text-xs font-bold text-slate-900 mb-1.5">{label}</label>
    <select
      value={value}
      onChange={e => onChange(e.target.value)}
      className="w-full bg-white border border-[#CBD5E1] rounded-lg px-3 py-2.5 text-sm text-slate-900 focus:outline-none focus:border-[#2563EB] focus:ring-2 focus:ring-blue-200 transition-all"
    >
      {options.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
    </select>
  </div>
);

const Btn = ({ children, onClick, variant = 'primary', size = 'md', icon: Icon, fullWidth, disabled }) => {
  const variants = {
    primary: 'text-white border-transparent font-bold',
    secondary: 'bg-white hover:bg-[#F1F5F9] border-[#D8E4F8] hover:border-[#BFDBFE] font-bold',
    danger: 'bg-white hover:bg-[#FFE4E6] border-[#FECDD3] font-bold',
    dangerSolid: 'bg-[#F43F5E] hover:bg-[#E11D48] text-white border-[#F43F5E] font-bold',
    ghost: 'bg-transparent border-transparent hover:bg-[#F1F5F9] font-bold'
  };
  const sizes = { sm: 'px-3 py-1.5 text-xs gap-1.5', md: 'px-4 py-2 text-sm gap-2', lg: 'px-5 py-2.5 text-sm gap-2' };
  const styleProps = variant === 'primary' ? {
    background: '#2563EB',
    borderRadius: '14px',
    boxShadow: '0 8px 20px rgba(37, 99, 235, 0.25)',
    color: '#FFFFFF'
  } : variant === 'secondary' ? {
    borderRadius: '14px',
    color: '#0F172A'
  } : variant === 'danger' ? {
    borderRadius: '14px',
    color: '#F43F5E'
  } : {
    borderRadius: '14px',
    color: '#475569'
  };
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      onMouseEnter={e => variant === 'primary' && !disabled && (e.currentTarget.style.background = '#1D4ED8')}
      onMouseLeave={e => variant === 'primary' && !disabled && (e.currentTarget.style.background = '#2563EB')}
      className={`${variants[variant]} ${sizes[size]} ${fullWidth ? 'w-full justify-center' : ''} border inline-flex items-center transition-all disabled:opacity-50 disabled:cursor-not-allowed`}
      style={styleProps}
    >
      {Icon && <Icon size={size === 'sm' ? 14 : 16} strokeWidth={2.4} />}
      {children}
    </button>
  );
};

// ============================================================
// TOOLTIP — works on hover (desktop) and tap (mobile)
// ============================================================
const Tooltip = ({ text, position = 'top', children }) => {
  const [open, setOpen] = useState(false);

  // Close on outside click for mobile
  useEffect(() => {
    if (!open) return;
    const handler = () => setOpen(false);
    setTimeout(() => document.addEventListener('click', handler), 0);
    return () => document.removeEventListener('click', handler);
  }, [open]);

  const positions = {
    top: 'bottom-full left-1/2 -translate-x-1/2 mb-2',
    bottom: 'top-full left-1/2 -translate-x-1/2 mt-2',
    left: 'right-full top-1/2 -translate-y-1/2 mr-2',
    right: 'left-full top-1/2 -translate-y-1/2 ml-2'
  };

  const arrowPositions = {
    top: 'top-full left-1/2 -translate-x-1/2 border-t-slate-900 border-l-transparent border-r-transparent border-b-transparent',
    bottom: 'bottom-full left-1/2 -translate-x-1/2 border-b-slate-900 border-l-transparent border-r-transparent border-t-transparent',
    left: 'left-full top-1/2 -translate-y-1/2 border-l-slate-900 border-t-transparent border-b-transparent border-r-transparent',
    right: 'right-full top-1/2 -translate-y-1/2 border-r-slate-900 border-t-transparent border-b-transparent border-l-transparent'
  };

  return (
    <span
      className="relative inline-flex group"
      onMouseEnter={() => setOpen(true)}
      onMouseLeave={() => setOpen(false)}
      onClick={(e) => { e.stopPropagation(); setOpen(o => !o); }}
    >
      {children}
      {open && (
        <span className={`absolute ${positions[position]} z-50 pointer-events-none`}>
          <span className="block bg-white text-slate-900 text-xs leading-relaxed rounded-lg px-3 py-2 shadow-xl w-64 whitespace-normal">
            {text}
          </span>
          <span className={`absolute w-0 h-0 border-4 ${arrowPositions[position]}`} />
        </span>
      )}
    </span>
  );
};

// Inline help icon next to labels
const InfoIcon = ({ text, position = 'top' }) => (
  <Tooltip text={text} position={position}>
    <HelpCircle size={12} className="text-blue-700 hover:text-slate-900 cursor-help inline-block ml-1 align-middle" strokeWidth={2} />
  </Tooltip>
);

// ============================================================
// GLOSSARY — central explanations for every term
// ============================================================
const GLOSSARY = {
  // KPIs
  totalProjects: 'The total count of all projects you have entered into the system, regardless of status.',
  contractValue: 'The agreed total amount the client will pay for this project, including VAT. The headline number on the contract.',
  projectedProfit: 'Your final profit after subtracting VAT, all costs, deductions, and partner payouts. This is the money you actually keep.',
  totalDeductions: 'Every cost subtracted from the project: bonds, insurance, mobilisation, custom expenses, partner payouts, etc. Does not include VAT.',
  realCash: 'The cash actually available to you upfront after the bank holds back retention and you pay out fixed costs. What you can use right now.',
  activeProjects: 'Projects with status set to "Active" — currently in progress, not yet completed or cancelled.',
  riskyProjects: 'Projects flagged as risky because they have negative cash, negative profit, or other red flags. Review before committing capital.',
  bestProject: 'The project with the highest projected final profit. Tap to jump straight to its detail page.',

  // Project metrics
  effectiveMargin: 'Your final profit as a percentage of contract value. A 20% margin on 1B contract = 200M profit. Higher is better.',
  capitalEfficiency: 'How many shillings of profit you earn for every shilling of cost/deduction. 3x means strong leverage; below 1x means costs eat your profit.',
  expectedProfit: 'Profit before deductions — calculated as profit margin % times net value (after VAT). The maximum theoretical profit.',
  advanceCash: 'The upfront payment from the client minus what the bank holds in retention. The actual cash that hits your account day one.',
  bankRetention: 'The percentage of advance the bank holds back as security, released only when project milestones are met.',
  vat: 'Value Added Tax charged on the contract. Kenya standard is 16%. This is collected on behalf of the government — not your money.',
  netAfterVat: 'Contract value minus VAT. The actual revenue figure used to calculate your profit margin.',
  advancePct: 'The percentage of total contract paid upfront. Higher advance = more working capital, less risk.',
  profitMargin: 'Your expected profit as a percentage of net (after-VAT) value. Industry varies: construction 8-15%, supply 5-10%, services 20-30%.',
  riskLevel: 'SAFE = all indicators healthy. CAUTION = at least one warning sign. RISKY = negative cash or profit detected.',
  healthScore: 'A 0-100 score combining cash position, profit, margin, capital efficiency, and coverage (bonds, insurance). 75+ is healthy, 50-74 needs attention, below 50 is critical.',

  // Form fields
  contractValueField: 'The total agreed contract value with the client. Enter the gross amount including VAT.',
  vatField: 'VAT percentage charged. Standard rate in Kenya is 16%. Some sectors are exempt or zero-rated.',
  profitField: 'Expected profit margin on net (after-VAT) value. Be realistic — do not pad. This drives all profit calculations.',
  advanceField: 'Percentage of contract paid upfront before work starts. Construction often 10-30%, supply often 30-50%.',
  retentionField: 'Percentage of advance the bank holds as guarantee. Released as project milestones are completed.',
  mobilisationField: 'Initial site setup costs: equipment hire, temporary structures, material delivery, workforce mobilisation.',
  bondField: 'Performance bond fee — a guarantee that the project will be completed. Usually 1-3% of contract value, paid to bank/insurer.',
  guaranteeField: 'Bank guarantee fee — protects the client if you default. Usually 1-2% of advance value.',
  insuranceField: 'Project insurance — covers theft, damage, accidents, third-party liability. Typically 0.5-2% of contract.',
  tenderField: 'Tender or application fees paid to bid for the project. Non-refundable in most cases.',
  legalField: 'Legal documentation, contract review, registrations, permits, and compliance costs.',
  commissionField: 'Commissions paid to brokers, agents, or facilitators who helped secure the contract.',
  customDeductionField: 'Project-specific costs not covered by standard categories. Add anything that subtracts from your profit.',
  partnerField: 'Co-investors or shareholders who get a percentage of profit OR a fixed payout.',
  partnerShareField: 'Percentage of final profit this partner receives. Computed after all deductions.',
  partnerPayoutField: 'Fixed amount paid to this partner regardless of profit. Subtracted from upfront cash.',

  // Comparison
  upfrontCash: 'The advance payment after bank retention is held back. The first cash you receive.',
  finalProfit: 'Net profit you keep after every deduction. The bottom-line number.',

  // Status
  planning: 'Project is still being evaluated, contract not yet signed, or work has not begun.',
  active: 'Contract signed, work in progress, currently consuming time and capital.',
  completed: 'Project finished, all payments received and disbursed.',
  cancelled: 'Project did not proceed. Kept in records for reference.',

  // Stress tests
  stressTests: 'Automated what-if scenarios showing how the project survives margin drops, expense spikes, and reduced advance. If profit stays positive across all four, the project is robust.'
};

const PageHeader = ({ title, subtitle, action }) => (
  <div className="flex flex-wrap items-start justify-between gap-4 mb-8">
    <div>
      <h1 className="font-display text-[34px] font-extrabold tracking-tight leading-tight" style={{ color: '#071739' }}>{title}</h1>
      {subtitle && <p className="text-sm mt-1.5 font-bold" style={{ color: '#2563EB' }}>{subtitle}</p>}
    </div>
    {action}
  </div>
);

const EmptyState = ({ icon: Icon, title, subtitle, action }) => (
  <Card padding="p-12">
    <div className="text-center">
      <div className="inline-flex items-center justify-center w-14 h-14 bg-[#EFF6FF] rounded-full mb-4">
        <Icon size={24} className="text-[#2563EB]" strokeWidth={1.5} />
      </div>
      <h3 className="text-lg font-bold text-slate-900 mb-1">{title}</h3>
      {subtitle && <p className="text-sm text-blue-700 mb-6 max-w-md mx-auto">{subtitle}</p>}
      {action}
    </div>
  </Card>
);

// ============================================================
// SIDEBAR
// ============================================================
const Sidebar = ({ view, setView, projectCount, mobileOpen, setMobileOpen, onLogout, onExport, onImport }) => {
  const items = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'projects', label: 'Projects', icon: FolderKanban, badge: projectCount },
    { id: 'compare', label: 'Compare', icon: GitCompare, disabled: projectCount < 2 },
    { id: 'reports', label: 'Reports', icon: FileText, disabled: projectCount === 0 }
  ];

  const handleClick = (id, disabled) => {
    if (disabled) return;
    setView(id);
    setMobileOpen(false);
  };

  return (
    <>
      {mobileOpen && (
        <div className="fixed inset-0 bg-white/50 z-40 lg:hidden" onClick={() => setMobileOpen(false)} />
      )}
      <aside className={`fixed lg:sticky top-0 left-0 h-screen w-64 bg-white text-slate-900 flex flex-col z-50 transform transition-transform duration-200 border-r ${mobileOpen ? 'translate-x-0' : '-translate-x-full lg:translate-x-0'}`} style={{ borderColor: '#E2E8F0' }}>
        <div className="px-5 py-6 border-b flex items-center gap-3" style={{ borderColor: '#E2E8F0' }}>
          <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{
            background: '#2563EB',
            boxShadow: '0 6px 16px rgba(37, 99, 235, 0.25)'
          }}>
            <Calculator size={20} className="text-white" strokeWidth={2.4} />
          </div>
          <div>
            <div className="font-display font-extrabold text-base tracking-tight" style={{ color: '#071739' }}>Rays InvestCalc</div>
            <div className="text-[10px] uppercase tracking-wider font-bold mt-0.5" style={{ color: '#2563EB' }}>Decision Engine</div>
          </div>
        </div>

        <nav className="flex-1 p-3 space-y-0.5">
          <div className="text-[10px] uppercase tracking-wider font-bold px-3 py-2.5" style={{ color: '#64748B' }}>Workspace</div>
          {items.map(item => {
            const active = view === item.id || (item.id === 'projects' && view === 'detail');
            return (
              <button
                key={item.id}
                onClick={() => handleClick(item.id, item.disabled)}
                disabled={item.disabled}
                className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-bold transition-all ${
                  active
                    ? 'text-white'
                    : item.disabled
                    ? 'cursor-not-allowed'
                    : ''
                }`}
                style={active ? {
                  background: '#2563EB',
                  boxShadow: '0 6px 16px rgba(37, 99, 235, 0.25)'
                } : item.disabled ? {
                  color: '#CBD5E1'
                } : {
                  color: '#475569'
                }}
                onMouseEnter={e => !active && !item.disabled && (e.currentTarget.style.background = '#F1F5F9')}
                onMouseLeave={e => !active && !item.disabled && (e.currentTarget.style.background = 'transparent')}
              >
                <item.icon size={16} strokeWidth={active ? 2.4 : 2} />
                <span className="flex-1 text-left">{item.label}</span>
                {item.badge !== undefined && item.badge > 0 && (
                  <span className="text-[10px] px-1.5 py-0.5 rounded-md font-bold" style={active ? { background: 'rgba(255,255,255,0.25)', color: '#FFFFFF' } : { background: '#DBEAFE', color: '#2563EB' }}>
                    {item.badge}
                  </span>
                )}
              </button>
            );
          })}
        </nav>

        <div className="p-4 border-t border-[#D8E4F8] space-y-2">
          <div className="bg-[#EFF6FF] border border-[#D8E4F8] rounded-lg p-3">
            <div className="flex items-center gap-2 mb-1.5">
              <div className="w-1.5 h-1.5 bg-[#EFF6FF]0 rounded-full animate-pulse" />
              <div className="text-[10px] font-bold text-slate-900 uppercase tracking-wider">Live Workspace</div>
            </div>
            <div className="text-[11px] text-[#697386] leading-relaxed">Track, analyze, and compare every project.</div>
          </div>

          <div className="grid grid-cols-2 gap-2">
            <button
              onClick={onExport}
              title="Export backup"
              className="flex flex-col items-center justify-center gap-1 px-2 py-2.5 text-[10px] font-bold text-slate-700 hover:text-blue-500 hover:bg-[#F1F5F9] rounded-lg transition-colors border border-[#D8E4F8]"
            >
              <Download size={13} strokeWidth={2} />
              Backup
            </button>
            <label
              title="Import backup"
              className="flex flex-col items-center justify-center gap-1 px-2 py-2.5 text-[10px] font-bold text-slate-700 hover:text-blue-500 hover:bg-[#F1F5F9] rounded-lg transition-colors border border-[#D8E4F8] cursor-pointer"
            >
              <Upload size={13} strokeWidth={2} />
              Restore
              <input type="file" accept=".json,application/json" onChange={onImport} className="hidden" />
            </label>
          </div>

          <button
            onClick={onLogout}
            className="w-full flex items-center justify-center gap-2 px-3 py-2.5 text-xs font-bold text-slate-700 hover:text-slate-900 hover:bg-[#F1F5F9] rounded-lg transition-colors border border-[#D8E4F8]"
          >
            <LogOut size={14} strokeWidth={2} />
            Lock & sign out
          </button>
        </div>
      </aside>
    </>
  );
};

// ============================================================
// DASHBOARD
// ============================================================
const Dashboard = ({ projects, onNew, onKpiClick, onView }) => {
  const stats = useMemo(() => {
    const calcs = projects.map(p => ({ p, c: calcProject(p) }));
    const active = calcs.filter(x => x.p.status === 'active');
    const totalContract = calcs.reduce((s, x) => s + x.c.contract, 0);
    const totalProfit = calcs.reduce((s, x) => s + x.c.projectedFinalProfit, 0);
    const totalDeductions = calcs.reduce((s, x) => s + x.c.totalDeductions, 0);
    const totalCash = calcs.reduce((s, x) => s + x.c.realCashAfterDeductions, 0);
    const sorted = [...calcs].sort((a, b) => b.c.projectedFinalProfit - a.c.projectedFinalProfit);
    const best = sorted[0];
    const risky = calcs.filter(x => x.c.riskLevel === 'risky');
    const caution = calcs.filter(x => x.c.riskLevel === 'caution');
    const avgMargin = calcs.length > 0 ? calcs.reduce((s, x) => s + x.c.profitMargin, 0) / calcs.length : 0;
    const portfolioEfficiency = totalDeductions > 0 ? totalProfit / totalDeductions : null;
    const negativeCash = calcs.filter(x => x.c.realCashAfterDeductions < 0);

    // Build insights
    const insights = [];
    if (risky.length > 0) {
      insights.push({ tone: 'rose', icon: ShieldAlert, title: `${risky.length} risky project${risky.length === 1 ? '' : 's'}`, body: `Review before committing capital: ${risky.slice(0, 2).map(x => x.p.name).join(', ')}${risky.length > 2 ? '…' : ''}` });
    }
    if (negativeCash.length > 0) {
      insights.push({ tone: 'amber', icon: AlertTriangle, title: 'Cash deficit on some projects', body: `${negativeCash.length} project${negativeCash.length === 1 ? ' will require' : 's will require'} out-of-pocket funding upfront.` });
    }
    if (best && best.c.profitMargin > 15) {
      insights.push({ tone: 'emerald', icon: Trophy, title: `Strongest margin: ${best.p.name}`, body: `Effective margin ${best.c.profitMargin.toFixed(1)}%${best.c.capitalEfficiency ? `, capital efficiency ${best.c.capitalEfficiency.toFixed(1)}x` : ''}.` });
    }
    if (caution.length > 0 && risky.length === 0) {
      insights.push({ tone: 'amber', icon: AlertTriangle, title: `${caution.length} project${caution.length === 1 ? '' : 's'} need${caution.length === 1 ? 's' : ''} attention`, body: `Caution flags raised — check deduction levels and advance terms.` });
    }
    if (portfolioEfficiency && portfolioEfficiency > 3) {
      insights.push({ tone: 'sky', icon: Sparkles, title: `Portfolio capital efficiency: ${portfolioEfficiency.toFixed(1)}x`, body: `Every shilling deployed returns ${portfolioEfficiency.toFixed(1)} in profit. Strong leverage.` });
    }

    // Concentration risk
    if (calcs.length >= 2) {
      const sorted = [...calcs].sort((a, b) => b.c.contract - a.c.contract);
      const concentration = totalContract > 0 ? (sorted[0].c.contract / totalContract) * 100 : 0;
      if (concentration > 70) {
        insights.push({ tone: 'amber', icon: ShieldAlert, title: 'Portfolio concentration risk', body: `${concentration.toFixed(0)}% of value in ${sorted[0].p.name}. One project drives most of your portfolio.` });
      }
    }

    // Single project portfolio
    if (calcs.length === 1) {
      insights.push({ tone: 'sky', icon: Lightbulb, title: 'Single project portfolio', body: 'Add more projects to compare opportunities and spread risk across deals.' });
    }

    // Currency mix
    const currencies = [...new Set(calcs.map(x => x.p.currency))];
    if (currencies.length > 1) {
      insights.push({ tone: 'sky', icon: DollarSign, title: 'Multi-currency exposure', body: `Operating in ${currencies.join(', ')}. Watch FX risk on cross-border projects.` });
    }

    // Status mix bottleneck
    const planning = calcs.filter(x => x.p.status === 'planning');
    if (planning.length >= 3 && active.length === 0) {
      insights.push({ tone: 'amber', icon: Activity, title: 'No active projects', body: `${planning.length} projects in planning. Pick one and move forward.` });
    }

    // High-margin opportunity
    const highMarginPlanning = planning.filter(x => x.c.profitMargin > 15 && x.c.riskLevel === 'safe');
    if (highMarginPlanning.length > 0 && active.length === 0) {
      insights.push({ tone: 'emerald', icon: Sparkles, title: 'Strong opportunity ready to launch', body: `${highMarginPlanning[0].p.name} has ${highMarginPlanning[0].c.profitMargin.toFixed(1)}% margin and is rated safe. Worth prioritizing.` });
    }

    // Total VAT exposure
    const totalVat = calcs.reduce((s, x) => s + x.c.vatAmount, 0);
    if (totalVat > totalProfit * 0.5) {
      insights.push({ tone: 'sky', icon: FileText, title: 'Significant VAT to manage', body: `${fmtShort(totalVat)} in VAT across portfolio. Stay current on filings to avoid penalties.` });
    }

    if (insights.length === 0 && calcs.length > 0) {
      insights.push({ tone: 'emerald', icon: CheckCircle2, title: 'Portfolio looks healthy', body: 'No major risk flags detected across your projects.' });
    }

    const recommendations = computeRecommendations(calcs);

    return { calcs, totalContract, totalProfit, totalDeductions, totalCash, best, risky, caution, active, avgMargin, portfolioEfficiency, insights, recommendations };
  }, [projects]);

  return (
    <div>
      <PageHeader
        title="Dashboard"
        subtitle="Overview of all your project investments"
        action={<Btn onClick={onNew} icon={Plus}>New Project</Btn>}
      />

      {projects.length > 0 && (() => {
        // Compute smart insights for action chips
        const allDocs = projects.flatMap(p => (p.documents || []).map(d => ({ ...d, projectId: p.id, projectName: p.name })));
        const expiringSoon = allDocs.filter(d => {
          if (!d.expiryDate || d.status !== 'received') return false;
          const days = Math.ceil((new Date(d.expiryDate) - new Date()) / (1000 * 60 * 60 * 24));
          return days >= 0 && days <= 30;
        }).length;
        const expired = allDocs.filter(d => {
          if (!d.expiryDate || d.status !== 'received') return false;
          return new Date(d.expiryDate) < new Date();
        }).length;
        const missingCritical = projects.reduce((sum, p) => {
          const ds = computeDocStatus(p);
          return sum + ds.missing.length;
        }, 0);
        const profitPositive = stats.totalProfit >= 0;
        const today = new Date();

        return (
          <div className="relative mb-6">
            {/* Premium portfolio snapshot hero - spec colors */}
            <div className="relative overflow-hidden" style={{
              background: 'linear-gradient(135deg, #F8FBFF 0%, #EEF6FF 55%, #DBEAFE 100%)',
              border: '1px solid #D8E4F8',
              borderRadius: '24px',
              boxShadow: '0 18px 45px rgba(37, 99, 235, 0.10)'
            }}>
              {/* Decorative financial pattern on the right */}
              <svg className="absolute right-0 top-0 h-full pointer-events-none" viewBox="0 0 480 360" preserveAspectRatio="xMaxYMid slice" style={{ width: '45%', opacity: 1 }}>
                <defs>
                  <linearGradient id="barGrad" x1="0%" y1="0%" x2="0%" y2="100%">
                    <stop offset="0%" stopColor="#2563EB" stopOpacity="0.18" />
                    <stop offset="100%" stopColor="#2563EB" stopOpacity="0.04" />
                  </linearGradient>
                  <linearGradient id="barGrad2" x1="0%" y1="0%" x2="0%" y2="100%">
                    <stop offset="0%" stopColor="#3B82F6" stopOpacity="0.12" />
                    <stop offset="100%" stopColor="#3B82F6" stopOpacity="0.03" />
                  </linearGradient>
                </defs>
                {/* Faint wave lines */}
                <path d="M 0 220 Q 80 200 160 180 T 320 150 T 560 120" stroke="rgba(37, 99, 235, 0.10)" strokeWidth="1.5" fill="none" />
                <path d="M 0 250 Q 80 230 160 210 T 320 180 T 560 150" stroke="rgba(37, 99, 235, 0.08)" strokeWidth="1.5" fill="none" />
                <path d="M 0 280 Q 80 260 160 240 T 320 210 T 560 180" stroke="rgba(37, 99, 235, 0.06)" strokeWidth="1.5" fill="none" />
                {/* Ascending bar chart */}
                <rect x="280" y="220" width="28" height="120" rx="5" fill="url(#barGrad2)" />
                <rect x="318" y="180" width="28" height="160" rx="5" fill="url(#barGrad2)" />
                <rect x="356" y="135" width="28" height="205" rx="5" fill="url(#barGrad)" />
                <rect x="394" y="90" width="28" height="250" rx="5" fill="url(#barGrad)" />
                <rect x="432" y="55" width="28" height="285" rx="5" fill="url(#barGrad)" />
              </svg>

              <div className="relative p-6 sm:p-8">
                {/* Eyebrow */}
                <div className="text-[11px] uppercase tracking-[0.16em] font-extrabold mb-3" style={{ color: '#2563EB' }}>
                  Portfolio Snapshot · {today.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' }).toUpperCase()}
                </div>

                {/* Big primary number + positive badge */}
                <div className="flex items-center gap-3 mb-2 flex-wrap">
                  <div className="num-display text-[44px] sm:text-[56px] leading-none" style={{ color: '#071739', letterSpacing: '-1.5px' }}>
                    {fmtShort(stats.totalProfit)}
                  </div>
                  <div className={`px-3 py-1.5 rounded-full text-[12px] font-extrabold flex items-center gap-1.5`} style={profitPositive ? {
                    background: '#DCFCE7',
                    color: '#16A34A',
                    border: '1px solid rgba(22, 163, 74, 0.20)'
                  } : {
                    background: '#FFE4E6',
                    color: '#F43F5E',
                    border: '1px solid rgba(244, 63, 94, 0.20)'
                  }}>
                    {profitPositive ? <TrendingUp size={13} strokeWidth={2.8} /> : <TrendingDown size={13} strokeWidth={2.8} />}
                    {profitPositive ? 'Positive' : 'Negative'}
                  </div>
                </div>
                <div className="text-[15px] font-semibold mb-6" style={{ color: '#334155' }}>
                  Projected profit across <span className="font-extrabold" style={{ color: '#071739' }}>{projects.length} {projects.length === 1 ? 'project' : 'projects'}</span>
                </div>

                {/* Metric strip - glass card with dividers */}
                <div className="grid grid-cols-2 sm:grid-cols-4 mb-6 overflow-hidden" style={{
                  background: 'rgba(255, 255, 255, 0.80)',
                  border: '1px solid #E2E8F0',
                  borderRadius: '20px',
                  backdropFilter: 'blur(8px)'
                }}>
                  <div className="p-5 sm:border-r" style={{ borderColor: '#D8E4F8' }}>
                    <div className="text-[10px] uppercase tracking-[0.12em] font-extrabold mb-1.5" style={{ color: '#475569' }}>Combined Contracts</div>
                    <div className="num-hero text-[22px] sm:text-[24px]" style={{ color: '#071739' }}>{fmtShort(stats.totalContract)}</div>
                  </div>
                  <div className="p-5 sm:border-r" style={{ borderColor: '#D8E4F8' }}>
                    <div className="text-[10px] uppercase tracking-[0.12em] font-extrabold mb-1.5" style={{ color: '#475569' }}>Real Cash Upfront</div>
                    <div className="num-hero text-[22px] sm:text-[24px]" style={{ color: '#071739' }}>{fmtShort(stats.totalCash)}</div>
                  </div>
                  <div className="p-5 sm:border-r border-t sm:border-t-0" style={{ borderColor: '#D8E4F8' }}>
                    <div className="text-[10px] uppercase tracking-[0.12em] font-extrabold mb-1.5" style={{ color: '#475569' }}>Avg Margin</div>
                    <div className="num-hero text-[22px] sm:text-[24px]" style={{ color: '#071739' }}>{stats.avgMargin.toFixed(1)}%</div>
                  </div>
                  <div className="p-5 border-t sm:border-t-0" style={{ borderColor: '#D8E4F8' }}>
                    <div className="text-[10px] uppercase tracking-[0.12em] font-extrabold mb-1.5" style={{ color: '#475569' }}>Capital Efficiency</div>
                    <div className="num-hero text-[22px] sm:text-[24px]" style={{ color: '#071739' }}>{stats.portfolioEfficiency ? stats.portfolioEfficiency.toFixed(2) + 'x' : '—'}</div>
                  </div>
                </div>

                {/* Two-column: needs attention | health distribution */}
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-6">
                  {/* LEFT: needs attention */}
                  <div>
                    <div className="text-[10px] uppercase tracking-[0.12em] font-extrabold mb-3" style={{ color: '#475569' }}>Needs Attention</div>
                    <div className="flex flex-wrap gap-2">
                      {stats.risky.length > 0 && (
                        <button onClick={() => onKpiClick('risky')} className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[11px] font-extrabold transition-all hover:scale-[1.02]" style={{ background: '#FFE4E6', color: '#F43F5E', border: '1px solid rgba(244, 63, 94, 0.20)' }}>
                          <AlertTriangle size={12} strokeWidth={2.6} />
                          {stats.risky.length} risky {stats.risky.length === 1 ? 'project' : 'projects'}
                        </button>
                      )}
                      {missingCritical > 0 && (
                        <button className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[11px] font-extrabold transition-all hover:scale-[1.02]" style={{ background: '#EFF6FF', color: '#334155', border: '1px solid #BFDBFE' }}>
                          <FileX size={12} strokeWidth={2.6} style={{ color: '#2563EB' }} />
                          {missingCritical} critical doc{missingCritical === 1 ? '' : 's'} missing
                        </button>
                      )}
                      {expired > 0 && (
                        <button className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[11px] font-extrabold transition-all hover:scale-[1.02]" style={{ background: '#FFE4E6', color: '#F43F5E', border: '1px solid rgba(244, 63, 94, 0.20)' }}>
                          <XCircle size={12} strokeWidth={2.6} />
                          {expired} expired
                        </button>
                      )}
                      {expiringSoon > 0 && (
                        <button className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[11px] font-extrabold transition-all hover:scale-[1.02]" style={{ background: '#FEF3C7', color: '#92400E', border: '1px solid rgba(245, 158, 11, 0.30)' }}>
                          <Clock size={12} strokeWidth={2.6} style={{ color: '#F59E0B' }} />
                          {expiringSoon} expiring soon
                        </button>
                      )}
                      {stats.risky.length === 0 && missingCritical === 0 && expired === 0 && expiringSoon === 0 && (
                        <div className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[11px] font-extrabold" style={{ background: '#DCFCE7', color: '#16A34A', border: '1px solid rgba(22, 163, 74, 0.20)' }}>
                          <CheckCircle2 size={12} strokeWidth={2.6} />
                          All projects on track
                        </div>
                      )}
                    </div>
                  </div>

                  {/* RIGHT: health distribution */}
                  <div className="lg:border-l lg:pl-8" style={{ borderColor: '#D8E4F8' }}>
                    <div className="text-[10px] uppercase tracking-[0.12em] font-extrabold mb-3" style={{ color: '#475569' }}>Health Distribution</div>
                    {(() => {
                      const total = projects.length || 1;
                      const safeCount = projects.filter(p => calcProject(p).riskLevel === 'safe').length;
                      const cautionCount = projects.filter(p => calcProject(p).riskLevel === 'caution').length;
                      const riskyCount = projects.filter(p => calcProject(p).riskLevel === 'risky').length;
                      const safePct = (safeCount / total) * 100;
                      const cautionPct = (cautionCount / total) * 100;
                      const riskyPct = (riskyCount / total) * 100;
                      return (
                        <>
                          <div className="flex h-2 rounded-full overflow-hidden mb-3" style={{ background: '#E2E8F0' }}>
                            {safePct > 0 && <div style={{ width: `${safePct}%`, background: '#2563EB' }} />}
                            {cautionPct > 0 && <div style={{ width: `${cautionPct}%`, background: '#F59E0B' }} />}
                            {riskyPct > 0 && <div style={{ width: `${riskyPct}%`, background: '#F43F5E' }} />}
                          </div>
                          <div className="grid grid-cols-3 gap-2">
                            <div className="flex items-center justify-between">
                              <div className="flex items-center gap-1.5">
                                <div className="w-2 h-2 rounded-full" style={{ background: '#2563EB' }} />
                                <span className="text-[12px] font-bold" style={{ color: '#0F172A' }}>Safe</span>
                              </div>
                              <span className="num-hero text-sm">{safeCount}</span>
                            </div>
                            <div className="flex items-center justify-between">
                              <div className="flex items-center gap-1.5">
                                <div className="w-2 h-2 rounded-full" style={{ background: '#F59E0B' }} />
                                <span className="text-[12px] font-bold" style={{ color: '#0F172A' }}>Caution</span>
                              </div>
                              <span className="num-hero text-sm">{cautionCount}</span>
                            </div>
                            <div className="flex items-center justify-between">
                              <div className="flex items-center gap-1.5">
                                <div className="w-2 h-2 rounded-full" style={{ background: '#F43F5E' }} />
                                <span className="text-[12px] font-bold" style={{ color: '#0F172A' }}>Risky</span>
                              </div>
                              <span className="num-hero text-sm">{riskyCount}</span>
                            </div>
                          </div>
                        </>
                      );
                    })()}
                  </div>
                </div>

                {/* Bottom row: Active | Total Locked */}
                <div className="grid grid-cols-2 gap-8 pt-6 border-t" style={{ borderColor: '#D8E4F8' }}>
                  <div>
                    <div className="text-[10px] uppercase tracking-[0.12em] font-extrabold mb-1.5" style={{ color: '#475569' }}>Active</div>
                    <div className="num-hero text-[24px]" style={{ color: '#071739' }}>{stats.active.length}<span className="text-base font-bold" style={{ color: '#94A3B8' }}>/{projects.length}</span></div>
                  </div>
                  <div>
                    <div className="text-[10px] uppercase tracking-[0.12em] font-extrabold mb-1.5" style={{ color: '#475569' }}>Total Locked</div>
                    <div className="num-hero text-[24px]" style={{ color: '#071739' }}>{fmtShort(stats.totalDeductions)}</div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-4">
        <KpiTile help={GLOSSARY.totalProjects} label="Total Projects" value={projects.length} sub="Across your portfolio" icon={FolderKanban} tone="slate" onClick={projects.length > 0 ? () => onKpiClick('totalProjects') : undefined} />
        <KpiTile help={GLOSSARY.contractValue} label="Contract Value" value={fmtShort(stats.totalContract)} sub="Combined total" icon={DollarSign} tone="emerald" onClick={projects.length > 0 ? () => onKpiClick('contract') : undefined} />
        <KpiTile help={GLOSSARY.projectedProfit} label="Projected Profit" value={fmtShort(stats.totalProfit)} sub="After all deductions" icon={TrendingUp} tone="teal" onClick={projects.length > 0 ? () => onKpiClick('profit') : undefined} />
        <KpiTile help={GLOSSARY.totalDeductions} label="Total Deductions" value={fmtShort(stats.totalDeductions)} sub="Costs & retentions" icon={TrendingDown} tone="amber" onClick={projects.length > 0 ? () => onKpiClick('deductions') : undefined} />
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
        <KpiTile help={GLOSSARY.realCash} label="Real Cash Available" value={fmtShort(stats.totalCash)} sub="Upfront after deductions" icon={Activity} tone="sky" onClick={projects.length > 0 ? () => onKpiClick('cash') : undefined} />
        <KpiTile help={GLOSSARY.activeProjects} label="Active Projects" value={stats.active.length} sub={`of ${projects.length} total`} icon={CheckCircle2} tone="indigo" onClick={stats.active.length > 0 ? () => onKpiClick('active') : undefined} />
        <KpiTile help={GLOSSARY.riskyProjects} label="Risky Projects" value={stats.risky.length} sub="Need review" icon={ShieldAlert} tone="rose" onClick={stats.risky.length > 0 ? () => onKpiClick('risky') : undefined} />
        <KpiTile help={GLOSSARY.bestProject} label="Best Project" value={stats.best?.p.name || 'N/A'} sub={stats.best ? fmtShort(stats.best.c.projectedFinalProfit) : 'Add projects'} icon={Trophy} tone="violet" onClick={stats.best ? () => onKpiClick('best', stats.best.p.id) : undefined} />
      </div>

      {projects.length === 0 ? (
        <EmptyState
          icon={FolderKanban}
          title="No projects yet"
          subtitle="Create your first project to start tracking investments and see real financial projections."
          action={<Btn onClick={onNew} icon={Plus}>Create Project</Btn>}
        />
      ) : (
        <>
          {/* Portfolio Charts */}
          {projects.length >= 1 && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-4">
              <Card padding="p-5">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <div className="flex items-center gap-2 mb-1">
                      <TrendingUp size={15} className="text-[#2563EB]" />
                      <h3 className="font-display font-bold text-[#0F172A]">{projects.length === 1 ? 'Profit Composition' : 'Profit by Project'}</h3>
                    </div>
                    <p className="text-xs text-[#64748B] font-medium">{projects.length === 1 ? 'How the contract becomes profit, KES millions' : 'Projected final profit, KES millions'}</p>
                  </div>
                </div>
                <div style={{ height: 240 }}>
                  {projects.length === 1 ? (() => {
                    const p = projects[0];
                    const c = calcProject(p);
                    const data = [
                      { name: 'Contract', value: c.contract / 1_000_000, fill: '#0F172A' },
                      { name: 'VAT', value: c.vatAmount / 1_000_000, fill: '#64748B' },
                      { name: 'Net', value: c.netAfterVat / 1_000_000, fill: '#2563EB' },
                      { name: 'Profit Pool', value: c.expectedProfit / 1_000_000, fill: '#3B82F6' },
                      { name: 'Deductions', value: c.totalDeductions / 1_000_000, fill: '#94A3B8' },
                      { name: 'Final Profit', value: c.projectedFinalProfit / 1_000_000, fill: c.projectedFinalProfit >= 0 ? '#16A34A' : '#F43F5E' }
                    ];
                    return (
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={data} margin={{ top: 10, right: 10, left: -10, bottom: 0 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                          <XAxis dataKey="name" tick={{ fill: '#64748B', fontSize: 10 }} stroke="#CBD5E1" />
                          <YAxis tick={{ fill: '#64748B', fontSize: 10 }} stroke="#CBD5E1" tickFormatter={(v) => `${v}M`} />
                          <RechartsTooltip
                            contentStyle={{ background: '#0F172A', border: 'none', color: '#fff', borderRadius: 8, fontSize: 11 }}
                            labelStyle={{ color: '#fff', fontWeight: 700 }}
                            cursor={{ fill: 'rgba(27,115,184,0.06)' }}
                            formatter={(v) => [`KES ${v.toFixed(1)}M`, '']}
                          />
                          <Bar dataKey="value" radius={[6, 6, 0, 0]}>
                            {data.map((d, i) => <Cell key={i} fill={d.fill} />)}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    );
                  })() : (
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart
                        data={projects.map(p => {
                          const c = calcProject(p);
                          return {
                            name: (p.name || 'Untitled').length > 14 ? (p.name || 'Untitled').slice(0, 12) + '…' : (p.name || 'Untitled'),
                            profit: Math.round(c.projectedFinalProfit / 1_000_000),
                            fill: c.projectedFinalProfit >= 0 ? '#16A34A' : '#F43F5E'
                          };
                        })}
                        margin={{ top: 10, right: 10, left: -10, bottom: 0 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                        <XAxis dataKey="name" tick={{ fill: '#64748B', fontSize: 10 }} stroke="#CBD5E1" />
                        <YAxis tick={{ fill: '#64748B', fontSize: 10 }} stroke="#CBD5E1" tickFormatter={(v) => `${v}M`} />
                        <RechartsTooltip
                          contentStyle={{ background: '#0F172A', border: 'none', color: '#fff', borderRadius: 8, fontSize: 11 }}
                          labelStyle={{ color: '#fff', fontWeight: 700 }}
                          cursor={{ fill: 'rgba(27,115,184,0.06)' }}
                          formatter={(v) => [`KES ${v.toLocaleString()}M`, 'Profit']}
                        />
                        <Bar dataKey="profit" radius={[6, 6, 0, 0]}>
                          {projects.map((p, i) => {
                            const c = calcProject(p);
                            return <Cell key={i} fill={c.projectedFinalProfit >= 0 ? '#16A34A' : '#F43F5E'} />;
                          })}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  )}
                </div>
                {projects.length === 1 && (() => {
                  const c = calcProject(projects[0]);
                  return (
                    <div className="grid grid-cols-3 gap-2 mt-3 pt-3 border-t border-[#E2E8F0]">
                      <div>
                        <div className="text-[9px] uppercase tracking-wider font-bold text-[#94A3B8] mb-0.5">Contract</div>
                        <div className="num-hero text-sm text-[#0F172A]">{fmtShort(c.contract)}</div>
                      </div>
                      <div>
                        <div className="text-[9px] uppercase tracking-wider font-bold text-[#94A3B8] mb-0.5">Deductions</div>
                        <div className="num-hero text-sm text-[#64748B]">{fmtShort(c.totalDeductions)}</div>
                      </div>
                      <div>
                        <div className="text-[9px] uppercase tracking-wider font-bold text-[#94A3B8] mb-0.5">Final Profit</div>
                        <div className={`num-hero text-sm ${c.projectedFinalProfit >= 0 ? 'text-[#2563EB]' : 'text-rose-600'}`}>{fmtShort(c.projectedFinalProfit)}</div>
                      </div>
                    </div>
                  );
                })()}
              </Card>

              <Card padding="p-5">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <div className="flex items-center gap-2 mb-1">
                      <Activity size={15} className="text-[#2563EB]" />
                      <h3 className="font-display font-bold text-slate-900">Cash Flow Composition</h3>
                    </div>
                    <p className="text-xs text-[#2563EB]">Where your capital is allocated</p>
                  </div>
                </div>
                {(() => {
                  const totalAdvance = projects.reduce((s, p) => s + calcProject(p).advanceAmount, 0);
                  const totalRetained = projects.reduce((s, p) => s + calcProject(p).bankRetained, 0);
                  const totalDeductions = projects.reduce((s, p) => s + calcProject(p).totalDeductions, 0);
                  const totalRealCash = projects.reduce((s, p) => s + calcProject(p).realCashAfterDeductions, 0);
                  const data = [
                    { name: 'Real Cash Deployable', value: Math.max(totalRealCash, 0), color: '#2563EB' },
                    { name: 'Bank Retention', value: totalRetained, color: '#0F172A' },
                    { name: 'Total Deductions', value: totalDeductions, color: '#94A3B8' }
                  ].filter(d => d.value > 0);
                  return (
                    <>
                      <div style={{ height: 200 }}>
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie
                              data={data}
                              cx="50%"
                              cy="50%"
                              innerRadius={55}
                              outerRadius={85}
                              paddingAngle={3}
                              dataKey="value"
                            >
                              {data.map((d, i) => <Cell key={i} fill={d.color} />)}
                            </Pie>
                            <RechartsTooltip
                              contentStyle={{ background: '#0F172A', border: 'none', color: '#fff', borderRadius: 8, fontSize: 11 }}
                              formatter={(v) => `KES ${(v/1_000_000).toFixed(1)}M`}
                            />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                      <div className="grid grid-cols-3 gap-2 mt-3">
                        {data.map((d, i) => (
                          <div key={i} className="text-center">
                            <div className="flex items-center justify-center gap-1.5 mb-1">
                              <div className="w-2 h-2 rounded-sm" style={{ background: d.color }} />
                              <span className="text-[9px] uppercase tracking-wider font-bold text-[#64748B]">{d.name}</span>
                            </div>
                            <div className="text-sm font-bold text-[#0F172A] tabular-nums">{fmtShort(d.value)}</div>
                          </div>
                        ))}
                      </div>
                    </>
                  );
                })()}
              </Card>
            </div>
          )}

          {/* Project Cash Arc — shows what happens from Day 1 to Final */}
          {projects.length >= 1 && (
            <Card className="mb-4" padding="p-5">
              <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <Activity size={15} className="text-[#2563EB]" />
                    <h3 className="font-display font-bold text-[#0F172A]">Cash Position Arc</h3>
                  </div>
                  <p className="text-xs text-[#64748B] font-medium">From contract value down to take-home profit (combined portfolio)</p>
                </div>
              </div>
              {(() => {
                const sum = projects.reduce((acc, p) => {
                  const c = calcProject(p);
                  acc.contract += c.contract;
                  acc.vat += c.vatAmount;
                  acc.profit += c.expectedProfit;
                  acc.advance += c.advanceAmount;
                  acc.retention += c.bankRetained;
                  acc.deductions += c.totalDeductions;
                  acc.realCash += c.realCashAfterDeductions;
                  acc.finalProfit += c.projectedFinalProfit;
                  return acc;
                }, { contract: 0, vat: 0, profit: 0, advance: 0, retention: 0, deductions: 0, realCash: 0, finalProfit: 0 });

                const arcData = [
                  { label: 'Contract', value: sum.contract / 1_000_000, color: '#0F172A' },
                  { label: 'VAT', value: sum.vat / 1_000_000, color: '#94A3B8' },
                  { label: 'Net Profit Pool', value: sum.profit / 1_000_000, color: '#3B82F6' },
                  { label: 'Advance', value: sum.advance / 1_000_000, color: '#2563EB' },
                  { label: 'Bank Held', value: sum.retention / 1_000_000, color: '#0F172A' },
                  { label: 'Deductions', value: sum.deductions / 1_000_000, color: '#CBD5E1' },
                  { label: 'Real Cash', value: sum.realCash / 1_000_000, color: '#2563EB' },
                  { label: 'Final Profit', value: sum.finalProfit / 1_000_000, color: sum.finalProfit >= 0 ? '#16A34A' : '#F43F5E' }
                ];
                return (
                  <div style={{ height: 220 }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={arcData} margin={{ top: 20, right: 10, left: -10, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                        <XAxis dataKey="label" tick={{ fill: '#64748B', fontSize: 10 }} stroke="#CBD5E1" />
                        <YAxis tick={{ fill: '#64748B', fontSize: 10 }} stroke="#CBD5E1" tickFormatter={(v) => `${v}M`} />
                        <RechartsTooltip
                          contentStyle={{ background: '#0F172A', border: 'none', color: '#fff', borderRadius: 8, fontSize: 11 }}
                          cursor={{ fill: 'rgba(27,115,184,0.06)' }}
                          formatter={(v) => [`KES ${v.toFixed(1)}M`, '']}
                        />
                        <Bar dataKey="value" radius={[6, 6, 0, 0]}>
                          {arcData.map((d, i) => <Cell key={i} fill={d.color} />)}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                );
              })()}
            </Card>
          )}

          {/* Insights Panel */}
          {stats.insights.length > 0 && (
            <Card className="mb-4" padding="p-5">
              <div className="flex items-center gap-2 mb-4">
                <div className="w-8 h-8 bg-gradient-to-br from-slate-400 to-slate-500 rounded-lg flex items-center justify-center">
                  <Lightbulb size={15} className="text-slate-900" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-900">Smart Insights</h3>
                  <p className="text-xs text-[#2563EB]">Auto-generated from your portfolio</p>
                </div>
              </div>
              <div className="space-y-2">
                {stats.insights.map((insight, i) => {
                  const tones = {
                    emerald: { bg: 'bg-[#EFF6FF]', border: 'border-[#CBD5E1]', text: 'text-blue-900', icon: 'text-blue-600' },
                    amber: { bg: 'bg-slate-50', border: 'border-slate-200', text: 'text-slate-900', icon: 'text-slate-600' },
                    rose: { bg: 'bg-rose-50', border: 'border-rose-200', text: 'text-rose-900', icon: 'text-rose-600' },
                    sky: { bg: 'bg-[#EFF6FF]', border: 'border-blue-200', text: 'text-blue-900', icon: 'text-blue-700' }
                  };
                  const t = tones[insight.tone];
                  return (
                    <div key={i} className={`${t.bg} border ${t.border} rounded-lg p-3 flex items-start gap-3`}>
                      <insight.icon size={16} className={`${t.icon} flex-shrink-0 mt-0.5`} strokeWidth={2} />
                      <div className="flex-1 min-w-0">
                        <div className={`text-sm font-bold ${t.text}`}>{insight.title}</div>
                        <div className={`text-xs ${t.text} opacity-90 mt-0.5`}>{insight.body}</div>
                      </div>
                    </div>
                  );
                })}
              </div>
              {stats.portfolioEfficiency !== null && (
                <div className="mt-4 pt-4 border-t border-[#EEF2F5] grid grid-cols-2 gap-4">
                  <div>
                    <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold">Portfolio Efficiency</div>
                    <div className="text-lg font-bold text-slate-900 tabular-nums mt-0.5">{stats.portfolioEfficiency.toFixed(2)}x</div>
                    <div className="text-[10px] text-[#2563EB]">Profit per shilling deployed</div>
                  </div>
                  <div>
                    <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold">Avg Margin</div>
                    <div className="text-lg font-bold text-slate-900 tabular-nums mt-0.5">{stats.avgMargin.toFixed(1)}%</div>
                    <div className="text-[10px] text-[#2563EB]">Across all projects</div>
                  </div>
                </div>
              )}
            </Card>
          )}

          {/* Recommendations Panel */}
          {stats.recommendations.length > 0 && (
            <Card className="mb-4" padding="p-5">
              <div className="flex items-center gap-2 mb-4">
                <div className="w-8 h-8 bg-gradient-to-br from-slate-500 to-slate-600 rounded-lg flex items-center justify-center">
                  <Sparkles size={15} className="text-slate-900" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-900">Recommendations</h3>
                  <p className="text-xs text-[#2563EB]">Actions that could improve your returns</p>
                </div>
              </div>
              <div className="space-y-2">
                {stats.recommendations.slice(0, 6).map((rec, i) => {
                  const priorityTones = {
                    high: { dot: 'bg-rose-500', label: 'HIGH', labelColor: 'text-rose-700 bg-rose-50 border-rose-200' },
                    medium: { dot: 'bg-slate-500', label: 'MED', labelColor: 'text-slate-700 bg-slate-50 border-slate-200' },
                    low: { dot: 'bg-slate-400', label: 'LOW', labelColor: 'text-blue-700 bg-[#EFF6FF] border-[#D8E4F8]' }
                  };
                  const pt = priorityTones[rec.priority];
                  const Icon = { DollarSign, Shield, AlertTriangle, Users, ShieldAlert, Activity, FileText, Lightbulb, Clock }[rec.icon] || Lightbulb;
                  const Element = rec.projectId ? 'button' : 'div';
                  return (
                    <Element
                      key={i}
                      onClick={rec.projectId ? () => onView(rec.projectId) : undefined}
                      className={`w-full text-left bg-white border border-[#CBD5E1] rounded-lg p-3 flex items-start gap-3 ${rec.projectId ? 'hover:border-slate-400 hover:shadow-sm cursor-pointer' : ''} transition-all group`}
                    >
                      <div className={`${pt.dot} w-1.5 h-1.5 rounded-full mt-2 flex-shrink-0`} />
                      <Icon size={15} className="text-blue-700 flex-shrink-0 mt-0.5" strokeWidth={2} />
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2 mb-0.5 flex-wrap">
                          <div className={`text-[9px] font-bold uppercase tracking-wider px-1.5 py-0.5 border rounded ${pt.labelColor}`}>{pt.label}</div>
                          <div className="text-sm font-bold text-slate-900 group-hover:text-slate-700 transition-colors">{rec.title}</div>
                        </div>
                        <div className="text-xs text-[#697386] leading-relaxed">{rec.body}</div>
                      </div>
                      {rec.projectId && <ChevronRight size={14} className="text-blue-700 group-hover:text-[#2563EB] mt-0.5 flex-shrink-0" />}
                    </Element>
                  );
                })}
                {stats.recommendations.length > 6 && (
                  <div className="text-center text-xs text-blue-700 pt-2">
                    + {stats.recommendations.length - 6} more recommendation{stats.recommendations.length - 6 === 1 ? '' : 's'}
                  </div>
                )}
              </div>
            </Card>
          )}

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          <Card className="lg:col-span-2" padding="p-5">
            <div className="flex items-center justify-between mb-5">
              <div>
                <h3 className="font-bold text-slate-900">Financial Overview</h3>
                <p className="text-xs text-blue-700 mt-0.5">Profit, cash, and deductions per project</p>
              </div>
            </div>
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={stats.calcs.slice(0, 8).map(x => ({
                name: (x.p.name || 'Untitled').length > 12 ? (x.p.name || 'Untitled').slice(0, 10) + '…' : (x.p.name || 'Untitled'),
                Profit: Math.round(x.c.projectedFinalProfit),
                Cash: Math.round(x.c.realCashAfterDeductions),
                Deductions: Math.round(x.c.totalDeductions)
              }))} margin={{ top: 5, right: 5, left: -10, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                <XAxis dataKey="name" stroke="#64748b" tick={{ fontSize: 11 }} axisLine={false} tickLine={false} />
                <YAxis stroke="#64748b" tick={{ fontSize: 11 }} tickFormatter={v => `${(v / 1000).toFixed(0)}k`} axisLine={false} tickLine={false} />
                <RechartsTooltip contentStyle={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: 8, fontSize: 12, boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                <Bar dataKey="Profit" fill="#1D4ED8" radius={[4, 4, 0, 0]} />
                <Bar dataKey="Cash" fill="#0284c7" radius={[4, 4, 0, 0]} />
                <Bar dataKey="Deductions" fill="#475569" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
            <div className="flex justify-center gap-6 mt-2">
              {[{ l: 'Profit', c: '#1D4ED8' }, { l: 'Cash', c: '#0284c7' }, { l: 'Deductions', c: '#475569' }].map(i => (
                <div key={i.l} className="flex items-center gap-2 text-xs text-[#2563EB]">
                  <div className="w-3 h-3 rounded-sm" style={{ background: i.c }} />
                  {i.l}
                </div>
              ))}
            </div>
          </Card>

          <Card padding="p-5">
            <h3 className="font-bold text-slate-900 mb-1">Risk Distribution</h3>
            <p className="text-xs text-blue-700 mb-4">Projects by risk level</p>
            {(() => {
              const pieData = [
                { name: 'Safe', value: stats.calcs.filter(x => x.c.riskLevel === 'safe').length, color: '#1D4ED8' },
                { name: 'Caution', value: stats.calcs.filter(x => x.c.riskLevel === 'caution').length, color: '#475569' },
                { name: 'Risky', value: stats.calcs.filter(x => x.c.riskLevel === 'risky').length, color: '#e11d48' }
              ].filter(x => x.value > 0);
              return (
                <>
                  <ResponsiveContainer width="100%" height={180}>
                    <PieChart>
                      <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={45} outerRadius={75} paddingAngle={2} stroke="white" strokeWidth={2}>
                        {pieData.map((entry, i) => <Cell key={i} fill={entry.color} />)}
                      </Pie>
                      <RechartsTooltip contentStyle={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: 8, fontSize: 12 }} />
                    </PieChart>
                  </ResponsiveContainer>
                  <div className="space-y-2 mt-3">
                    {pieData.map(x => (
                      <div key={x.name} className="flex items-center justify-between text-sm">
                        <div className="flex items-center gap-2">
                          <div className="w-2.5 h-2.5 rounded-full" style={{ background: x.color }} />
                          <span className="text-[#2563EB]">{x.name}</span>
                        </div>
                        <span className="font-bold text-slate-900">{x.value}</span>
                      </div>
                    ))}
                  </div>
                </>
              );
            })()}
          </Card>
        </div>
        </>
      )}
    </div>
  );
};

// ============================================================
// PROJECTS LIST
// ============================================================
const ProjectsList = ({ projects, onView, onNew }) => {
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');

  const filtered = useMemo(() => {
    return projects
      .map(p => ({ p, c: calcProject(p) }))
      .filter(({ p }) => {
        const q = search.toLowerCase();
        const matchSearch = !q || [p.name, p.client, p.location, p.type].some(f => (f || '').toLowerCase().includes(q));
        const matchStatus = statusFilter === 'all' || p.status === statusFilter;
        return matchSearch && matchStatus;
      });
  }, [projects, search, statusFilter]);

  return (
    <div>
      <PageHeader
        title="Projects"
        subtitle={`${projects.length} project${projects.length === 1 ? '' : 's'} in your portfolio`}
        action={<Btn onClick={onNew} icon={Plus}>New Project</Btn>}
      />

      <Card padding="p-4" className="mb-5">
        <div className="flex flex-col sm:flex-row gap-3">
          <div className="flex-1 relative">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-[#2563EB]" />
            <input
              type="text"
              placeholder="Search by name, client, location..."
              value={search}
              onChange={e => setSearch(e.target.value)}
              className="w-full pl-9 pr-3 py-2.5 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-[#2563EB] focus:ring-2 focus:ring-blue-200"
            />
          </div>
          <div className="relative">
            <Filter size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-blue-700 pointer-events-none" />
            <select
              value={statusFilter}
              onChange={e => setStatusFilter(e.target.value)}
              className="pl-9 pr-8 py-2.5 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-[#2563EB] appearance-none cursor-pointer"
            >
              <option value="all">All Statuses</option>
              <option value="planning">Planning</option>
              <option value="active">Active</option>
              <option value="completed">Completed</option>
              <option value="cancelled">Cancelled</option>
            </select>
          </div>
        </div>
      </Card>

      {filtered.length === 0 ? (
        <EmptyState
          icon={FolderKanban}
          title={projects.length === 0 ? 'No projects yet' : 'No projects found'}
          subtitle={projects.length === 0 ? 'Create your first project to get started.' : 'Try adjusting your search or filters.'}
          action={projects.length === 0 ? <Btn onClick={onNew} icon={Plus}>Create Project</Btn> : null}
        />
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
          {filtered.map(({ p, c }) => (
            <button
              key={p.id}
              onClick={() => onView(p.id)}
              className="text-left bg-white border border-[#CBD5E1]/80 rounded-2xl p-5 lift hover:border-[#2563EB]/30 transition-colors group relative overflow-hidden"
            >
              <div className="absolute top-0 right-0 w-32 h-32 bg-gradient-to-br from-blue-50 to-transparent rounded-full -translate-y-16 translate-x-16 opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none" />

              <div className="flex items-start justify-between mb-4 relative">
                <div className="flex-1 min-w-0">
                  <h3 className="font-display font-bold text-slate-900 truncate mb-1 group-hover:text-[#2563EB] transition-colors">{p.name || 'Untitled'}</h3>
                  <p className="text-xs text-blue-700 truncate">{p.client || 'No client'}</p>
                </div>
                <ChevronRight size={18} className="text-blue-700 group-hover:text-[#2563EB] group-hover:translate-x-0.5 transition-all flex-shrink-0" />
              </div>

              <div className="flex items-center gap-1.5 mb-4 relative">
                <Pill tone={p.status} size="sm">{p.status}</Pill>
                <Pill tone={c.riskLevel} size="sm">
                  {c.riskLevel === 'safe' && <CheckCircle2 size={10} />}
                  {c.riskLevel === 'caution' && <AlertTriangle size={10} />}
                  {c.riskLevel === 'risky' && <XCircle size={10} />}
                  {c.riskLevel}
                </Pill>
              </div>

              {p.location && (
                <div className="flex items-center gap-1.5 text-xs text-blue-700 mb-3 relative">
                  <MapPin size={12} />
                  <span className="truncate">{p.location}</span>
                </div>
              )}

              <div className="grid grid-cols-2 gap-3 pt-4 border-t border-[#EEF2F5] relative">
                <div>
                  <div className="text-[9px] uppercase tracking-[0.14em] text-[#2563EB] font-semibold mb-1">Contract</div>
                  <div className="num-hero text-base text-slate-900">{fmtShort(c.contract, p.currency)}</div>
                </div>
                <div>
                  <div className="text-[9px] uppercase tracking-[0.14em] text-[#2563EB] font-semibold mb-1">Profit</div>
                  <div className={`num-hero text-base ${c.projectedFinalProfit >= 0 ? 'text-emerald-700' : 'text-rose-600'}`}>
                    {fmtShort(c.projectedFinalProfit, p.currency)}
                  </div>
                </div>
              </div>
            </button>
          ))}
        </div>
      )}
    </div>
  );
};

// ============================================================
// KPI DRILLDOWN
// ============================================================
const Drilldown = ({ metric, projects, onBack, onView }) => {
  const data = projects.map(p => ({ p, c: calcProject(p) }));

  const configs = {
    contract: {
      title: 'Total Contract Value', subtitle: 'Combined contract value across all projects',
      icon: DollarSign, tone: 'emerald',
      getValue: x => x.c.contract, filter: () => true,
      getBreakdown: x => [
        { label: 'Contract value', v: x.c.contract, emphasis: true }
      ]
    },
    profit: {
      title: 'Projected Profit', subtitle: 'Final profit after all deductions',
      icon: TrendingUp, tone: 'teal',
      getValue: x => x.c.projectedFinalProfit, filter: () => true,
      getBreakdown: x => [
        { label: 'Net after VAT', v: x.c.netAfterVat },
        { label: `Expected profit (${fmtPct(x.p.profitPct)})`, v: x.c.expectedProfit, tone: 'amber' },
        { label: 'Less: Total deductions', v: -x.c.totalDeductions, tone: 'rose' },
        { label: 'Final profit', v: x.c.projectedFinalProfit, emphasis: true, tone: x.c.projectedFinalProfit >= 0 ? 'emerald' : 'rose' }
      ]
    },
    deductions: {
      title: 'Total Deductions', subtitle: 'All costs, fees & payouts across portfolio',
      icon: TrendingDown, tone: 'amber',
      getValue: x => x.c.totalDeductions, filter: () => true,
      getBreakdown: x => {
        const items = [
          { label: 'Mobilisation', v: +x.p.mobilisationCost || 0 },
          { label: 'Performance Bond', v: +x.p.performanceBondCost || 0 },
          { label: 'Bank Guarantee', v: +x.p.bankGuaranteeCost || 0 },
          { label: 'Insurance', v: +x.p.insuranceCost || 0 },
          { label: 'Tender / Application', v: +x.p.tenderCost || 0 },
          { label: 'Legal / Documentation', v: +x.p.legalCost || 0 },
          { label: 'Commissions', v: +x.p.commissionCost || 0 },
          ...(x.p.customDeductions || []).map(d => ({ label: d.label || 'Custom', v: +d.amount || 0 })),
          ...(x.p.partners || []).filter(pt => +pt.payout > 0).map(pt => ({ label: `Payout — ${pt.name || 'Partner'}`, v: +pt.payout || 0 }))
        ].filter(i => i.v > 0);
        items.push({ label: 'Total deductions', v: x.c.totalDeductions, emphasis: true, tone: 'rose' });
        return items;
      }
    },
    cash: {
      title: 'Real Cash Available', subtitle: 'Upfront cash after all deductions',
      icon: Activity, tone: 'sky',
      getValue: x => x.c.realCashAfterDeductions, filter: () => true,
      getBreakdown: x => [
        { label: `Advance (${fmtPct(x.p.advancePct)})`, v: x.c.advanceAmount },
        { label: `Bank retention (${fmtPct(x.p.bankRetentionPct)})`, v: -x.c.bankRetained, tone: 'rose' },
        { label: 'Advance cash', v: x.c.advanceCashAvailable, tone: 'amber' },
        { label: 'Less: Total deductions', v: -x.c.totalDeductions, tone: 'rose' },
        { label: 'Real cash upfront', v: x.c.realCashAfterDeductions, emphasis: true, tone: x.c.realCashAfterDeductions >= 0 ? 'emerald' : 'rose' }
      ]
    },
    active: {
      title: 'Active Projects', subtitle: 'Projects currently in progress',
      icon: CheckCircle2, tone: 'indigo',
      getValue: x => x.c.contract, filter: x => x.p.status === 'active', countMode: true
    },
    risky: {
      title: 'Risky Projects', subtitle: 'Projects flagged for review',
      icon: ShieldAlert, tone: 'rose',
      getValue: x => x.c.projectedFinalProfit, filter: x => x.c.riskLevel === 'risky', countMode: true
    }
  };

  const config = configs[metric];
  if (!config) return null;

  const filtered = data.filter(config.filter).sort((a, b) => Math.abs(config.getValue(b)) - Math.abs(config.getValue(a)));
  const total = filtered.reduce((s, x) => s + config.getValue(x), 0);
  const maxAbs = Math.max(...filtered.map(x => Math.abs(config.getValue(x))), 1);

  return (
    <div>
      <div className="flex items-center gap-3 mb-6">
        <button onClick={onBack} className="p-2 hover:bg-[#F1F5F9] rounded-lg text-blue-700 transition-colors">
          <ArrowLeft size={18} />
        </button>
        <div>
          <h1 className="font-display text-3xl font-extrabold text-slate-900">{config.title}</h1>
          <p className="text-sm text-blue-700 mt-0.5">{config.subtitle}</p>
        </div>
      </div>

      <div className="mb-6">
        <KpiTile
          label={config.countMode ? `Project Count` : 'Combined Total'}
          value={config.countMode ? filtered.length : fmtShort(total)}
          sub={config.countMode ? `${fmtShort(total)} combined value` : `Across ${filtered.length} project${filtered.length === 1 ? '' : 's'}`}
          icon={config.icon}
          tone={config.tone}
        />
      </div>

      {filtered.length === 0 ? (
        <EmptyState icon={config.icon} title="Nothing to show" subtitle="No projects match this metric." />
      ) : (
        <Card padding="p-5">
          <h3 className="font-bold text-slate-900 mb-1">Breakdown by Project</h3>
          <p className="text-xs text-blue-700 mb-5">Tap any project to see its full calculation</p>
          <div className="space-y-4">
            {filtered.map(({ p, c }) => {
              const value = config.getValue({ p, c });
              const pct = maxAbs > 0 ? (Math.abs(value) / maxAbs) * 100 : 0;
              const totalPct = total !== 0 ? (value / total) * 100 : 0;
              const breakdown = config.getBreakdown ? config.getBreakdown({ p, c }) : null;

              return (
                <button
                  key={p.id}
                  onClick={() => onView(p.id)}
                  className="w-full text-left bg-white border border-[#CBD5E1] rounded-lg p-4 hover:border-[#2563EB]/50 hover:shadow-sm transition-all group"
                >
                  <div className="flex items-start justify-between mb-3 gap-3">
                    <div className="min-w-0 flex-1">
                      <div className="font-bold text-slate-900 truncate group-hover:text-[#2563EB] transition-colors">{p.name || 'Untitled'}</div>
                      <div className="flex items-center gap-1.5 mt-1.5 flex-wrap">
                        <Pill tone={p.status} size="sm">{p.status}</Pill>
                        <Pill tone={c.riskLevel} size="sm">{c.riskLevel}</Pill>
                        {p.location && <span className="text-xs text-[#2563EB]">{p.location}</span>}
                      </div>
                    </div>
                    <div className="text-right flex-shrink-0">
                      <div className={`text-lg font-bold tabular-nums ${value >= 0 ? 'text-slate-900' : 'text-rose-600'}`}>
                        {fmtShort(value, p.currency)}
                      </div>
                      {!config.countMode && Math.abs(totalPct) < 1000 && (
                        <div className="text-xs text-blue-700 tabular-nums">{totalPct.toFixed(1)}% of total</div>
                      )}
                    </div>
                  </div>

                  <div className="h-1.5 bg-[#EFF6FF] rounded-full overflow-hidden mb-3">
                    <div
                      className={`h-full rounded-full ${value >= 0 ? 'bg-[#EFF6FF]0' : 'bg-rose-500'}`}
                      style={{ width: `${pct}%` }}
                    />
                  </div>

                  {breakdown && (
                    <div className="bg-[#EFF6FF] rounded-lg p-3 mt-3 space-y-1.5">
                      {breakdown.map((item, i) => {
                        const tones = { rose: 'text-rose-600', emerald: 'text-blue-600', amber: 'text-slate-600' };
                        return (
                          <div key={i} className={`flex justify-between items-baseline text-xs ${item.emphasis ? 'pt-1.5 border-t border-[#D8E4F8] mt-1.5' : ''}`}>
                            <span className={`${item.emphasis ? 'font-bold text-slate-900' : 'text-blue-700'}`}>{item.label}</span>
                            <span className={`font-mono tabular-nums ${tones[item.tone] || 'text-slate-900'} ${item.emphasis ? 'font-bold' : ''}`}>
                              {item.v < 0 ? '−' : ''}{fmt(Math.abs(item.v), { currency: p.currency })}
                            </span>
                          </div>
                        );
                      })}
                    </div>
                  )}

                  <div className="flex items-center justify-end mt-3 text-xs text-blue-600 opacity-0 group-hover:opacity-100 transition-opacity">
                    View full project <ChevronRight size={13} />
                  </div>
                </button>
              );
            })}
          </div>
        </Card>
      )}
    </div>
  );
};

// ============================================================
// PROJECT FORM
// ============================================================
const ProjectForm = ({ project, onSave, onCancel }) => {
  const [p, setP] = useState(project || blankProject());
  const upd = (k, v) => setP({ ...p, [k]: v });
  const c = calcProject(p);

  const addCustomDeduction = () => upd('customDeductions', [...(p.customDeductions || []), { id: Date.now(), label: '', amount: 0 }]);
  const updCustom = (i, k, v) => { const arr = [...p.customDeductions]; arr[i] = { ...arr[i], [k]: v }; upd('customDeductions', arr); };
  const rmCustom = i => upd('customDeductions', p.customDeductions.filter((_, idx) => idx !== i));

  const addPartner = () => upd('partners', [...(p.partners || []), { id: Date.now(), name: '', sharePct: 0, payout: 0 }]);
  const updPartner = (i, k, v) => { const arr = [...p.partners]; arr[i] = { ...arr[i], [k]: v }; upd('partners', arr); };
  const rmPartner = i => upd('partners', p.partners.filter((_, idx) => idx !== i));

  return (
    <div>
      <div className="flex flex-wrap items-start justify-between gap-4 mb-6">
        <div className="flex items-center gap-3">
          <button onClick={onCancel} className="p-2 hover:bg-[#F1F5F9] rounded-lg text-blue-700 transition-colors">
            <ArrowLeft size={18} />
          </button>
          <div>
            <h1 className="font-display text-3xl font-extrabold text-slate-900">{project ? 'Edit Project' : 'New Project'}</h1>
            <p className="text-sm text-blue-700 mt-0.5">Enter details to calculate real financial impact</p>
          </div>
        </div>
        <div className="flex gap-2">
          <Btn onClick={onCancel} variant="secondary">Cancel</Btn>
          <Btn onClick={() => onSave(p)} icon={Save}>Save Project</Btn>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2 space-y-5">
          <Card>
            <div className="flex items-center gap-2 mb-5">
              <div className="w-8 h-8 bg-[#DBEAFE] rounded-lg flex items-center justify-center">
                <Building2 size={15} className="text-[#2563EB]" />
              </div>
              <h2 className="font-bold text-slate-900">Project Details</h2>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <Input label="Project Name" value={p.name} onChange={v => upd('name', v)} required placeholder="e.g. Downtown Tower Construction" />
              <Input label="Client / Company" value={p.client} onChange={v => upd('client', v)} placeholder="e.g. ABC Corp" />
              <Input label="Location" value={p.location} onChange={v => upd('location', v)} placeholder="e.g. Nairobi, Kenya" />
              <Input label="Project Type" value={p.type} onChange={v => upd('type', v)} placeholder="e.g. Construction, Supply" />
              <Input label="Start Date" value={p.startDate} onChange={v => upd('startDate', v)} type="date" />
              <Input label="Duration (months)" value={p.duration} onChange={v => upd('duration', v)} placeholder="e.g. 6" />
              <Select label="Status" value={p.status} onChange={v => upd('status', v)} options={[
                { value: 'planning', label: 'Planning' },
                { value: 'active', label: 'Active' },
                { value: 'completed', label: 'Completed' },
                { value: 'cancelled', label: 'Cancelled' }
              ]} />
              <Select label="Currency" value={p.currency} onChange={v => upd('currency', v)} options={[
                { value: 'KES', label: 'KES — Kenyan Shilling' },
                { value: 'AUD', label: 'AUD — Australian Dollar' },
                { value: 'USD', label: 'USD — US Dollar' }
              ]} />
            </div>
          </Card>

          <Card>
            <div className="flex items-center gap-2 mb-5">
              <div className="w-8 h-8 bg-[#DBEAFE] rounded-lg flex items-center justify-center">
                <DollarSign size={15} className="text-[#2563EB]" />
              </div>
              <h2 className="font-bold text-slate-900">Financial Inputs</h2>
            </div>
            <div className="bg-[#EFF6FF] border border-blue-200 rounded-lg px-3 py-2 mb-4 flex items-start gap-2">
              <Info size={14} className="text-blue-700 flex-shrink-0 mt-0.5" />
              <p className="text-xs text-blue-900 leading-relaxed">
                Tap <span className="font-bold">%</span> or <span className="font-bold">{p.currency}</span> on each field to switch between entering a percentage or a fixed amount. The other value updates automatically.
              </p>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <Input help={GLOSSARY.contractValueField} label="Total Contract Value" value={p.contractValue} onChange={v => upd('contractValue', v)} type="number" prefix={p.currency} required />
              <PctAmountInput help={GLOSSARY.vatField} label="VAT" value={p.vatPct} onChange={v => upd('vatPct', v)} base={c.contract} currency={p.currency} hint="Kenya standard: 16%" baseLabel="of contract" storeAs="pct" />
              <PctAmountInput help={GLOSSARY.profitField} label="Expected Profit Margin" value={p.profitPct} onChange={v => upd('profitPct', v)} base={c.netAfterVat} currency={p.currency} hint="On net value after VAT" baseLabel="of net" storeAs="pct" />
              <PctAmountInput help={GLOSSARY.advanceField} label="Advance Payment" value={p.advancePct} onChange={v => upd('advancePct', v)} base={c.contract} currency={p.currency} hint="% of contract paid upfront" baseLabel="of contract" storeAs="pct" />
              <PctAmountInput help={GLOSSARY.retentionField} label="Bank Retention" value={p.bankRetentionPct} onChange={v => upd('bankRetentionPct', v)} base={c.advanceAmount} currency={p.currency} hint="% of advance held by bank" baseLabel="of advance" storeAs="pct" />
            </div>
          </Card>

          <Card>
            <div className="flex items-center gap-2 mb-5">
              <div className="w-8 h-8 bg-[#F1F5F9] rounded-lg flex items-center justify-center">
                <TrendingDown size={15} className="text-slate-600" />
              </div>
              <h2 className="font-bold text-slate-900">Costs & Deductions</h2>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <PctAmountInput help={GLOSSARY.mobilisationField} label="Mobilisation Cost" value={p.mobilisationCost} onChange={v => upd('mobilisationCost', v)} base={c.contract} currency={p.currency} hint="Typically 5-10% of contract" baseLabel="of contract" storeAs="amount" />
              <PctAmountInput help={GLOSSARY.bondField} label="Performance Bond" value={p.performanceBondCost} onChange={v => upd('performanceBondCost', v)} base={c.contract} currency={p.currency} hint="Typically 1-3% of contract" baseLabel="of contract" storeAs="amount" />
              <PctAmountInput help={GLOSSARY.guaranteeField} label="Bank Guarantee" value={p.bankGuaranteeCost} onChange={v => upd('bankGuaranteeCost', v)} base={c.advanceAmount} currency={p.currency} hint="Typically 1-2% of advance" baseLabel="of advance" storeAs="amount" />
              <PctAmountInput help={GLOSSARY.insuranceField} label="Insurance" value={p.insuranceCost} onChange={v => upd('insuranceCost', v)} base={c.contract} currency={p.currency} hint="Typically 0.5-2% of contract" baseLabel="of contract" storeAs="amount" />
              <PctAmountInput help={GLOSSARY.tenderField} label="Tender / Application" value={p.tenderCost} onChange={v => upd('tenderCost', v)} base={c.contract} currency={p.currency} hint="Application or bid fees" baseLabel="of contract" storeAs="amount" />
              <PctAmountInput help={GLOSSARY.legalField} label="Legal / Documentation" value={p.legalCost} onChange={v => upd('legalCost', v)} base={c.contract} currency={p.currency} hint="Contracts, permits, compliance" baseLabel="of contract" storeAs="amount" />
              <PctAmountInput help={GLOSSARY.commissionField} label="Commissions" value={p.commissionCost} onChange={v => upd('commissionCost', v)} base={c.contract} currency={p.currency} hint="Brokers, agents, facilitators" baseLabel="of contract" storeAs="amount" />
            </div>
          </Card>

          <Card>
            <div className="flex items-center justify-between mb-5">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 bg-[#F1F5F9] rounded-lg flex items-center justify-center">
                  <StickyNote size={15} className="text-[#2563EB]" />
                </div>
                <h2 className="font-bold text-slate-900">Custom Deductions</h2>
              </div>
              <Btn onClick={addCustomDeduction} variant="secondary" size="sm" icon={Plus}>Add</Btn>
            </div>
            {(p.customDeductions || []).length === 0 ? (
              <div className="text-sm text-blue-700 italic py-6 text-center border-2 border-dashed border-[#D8E4F8] rounded-lg">
                No custom deductions added yet
              </div>
            ) : (
              <div className="space-y-3">
                {p.customDeductions.map((d, i) => (
                  <div key={d.id || i} className="flex gap-2 items-end">
                    <div className="flex-1"><Input label="Description" value={d.label} onChange={v => updCustom(i, 'label', v)} /></div>
                    <div className="w-40"><Input label="Amount" value={d.amount} onChange={v => updCustom(i, 'amount', v)} type="number" prefix={p.currency} /></div>
                    <button onClick={() => rmCustom(i)} className="px-3 py-2.5 bg-white border border-rose-300 text-rose-600 hover:bg-rose-50 rounded-lg transition-colors">
                      <Trash2 size={15} />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </Card>

          <Card>
            <div className="flex items-center justify-between mb-5">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 bg-blue-50 rounded-lg flex items-center justify-center">
                  <Users size={15} className="text-blue-600" />
                </div>
                <h2 className="font-bold text-slate-900">Partners / Shareholders</h2>
              </div>
              <Btn onClick={addPartner} variant="secondary" size="sm" icon={Plus}>Add Partner</Btn>
            </div>
            {(p.partners || []).length === 0 ? (
              <div className="text-sm text-blue-700 italic py-6 text-center border-2 border-dashed border-[#D8E4F8] rounded-lg">
                No partners added yet
              </div>
            ) : (
              <div className="space-y-3">
                {p.partners.map((pt, i) => (
                  <div key={pt.id || i} className="grid grid-cols-12 gap-2 items-end">
                    <div className="col-span-12 sm:col-span-5"><Input label="Partner Name" value={pt.name} onChange={v => updPartner(i, 'name', v)} /></div>
                    <div className="col-span-6 sm:col-span-3"><Input label="Share" value={pt.sharePct} onChange={v => updPartner(i, 'sharePct', v)} type="number" suffix="%" /></div>
                    <div className="col-span-5 sm:col-span-3"><Input label="Direct Payout" value={pt.payout} onChange={v => updPartner(i, 'payout', v)} type="number" prefix={p.currency} /></div>
                    <div className="col-span-1"><button onClick={() => rmPartner(i)} className="w-full px-3 py-2.5 bg-white border border-rose-300 text-rose-600 hover:bg-rose-50 rounded-lg transition-colors"><Trash2 size={15} className="mx-auto" /></button></div>
                  </div>
                ))}
              </div>
            )}
          </Card>

          <Card>
            <div className="flex items-center gap-2 mb-4">
              <div className="w-8 h-8 bg-[#EFF6FF]0/15 rounded-lg ring-1 ring-blue-200 flex items-center justify-center">
                <StickyNote size={15} className="text-[#2563EB]" />
              </div>
              <h2 className="font-bold text-slate-900">Notes</h2>
            </div>
            <textarea
              value={p.notes}
              onChange={e => upd('notes', e.target.value)}
              rows={4}
              placeholder="Context, risks, decisions to remember..."
              className="w-full bg-white border border-slate-300 rounded-lg px-3 py-2.5 text-sm text-slate-900 placeholder:text-blue-700 focus:outline-none focus:border-[#2563EB] focus:ring-2 focus:ring-blue-200 resize-none"
            />
          </Card>

          {/* Documents */}
          <Card>
            <div className="flex items-center gap-2 mb-4">
              <div className="w-8 h-8 bg-[#DBEAFE] rounded-lg flex items-center justify-center">
                <FileCheck size={15} className="text-[#2563EB]" />
              </div>
              <div>
                <h2 className="font-bold text-slate-900">Documents</h2>
                <p className="text-xs text-[#2563EB]">Track contracts, bonds, insurance, permits, and partner agreements</p>
              </div>
            </div>
            <DocumentChecklist documents={p.documents} onChange={v => upd('documents', v)} projectId={p.id} />
          </Card>
        </div>

        <div className="lg:col-span-1">
          <div className="sticky top-4 space-y-4">
            <div className="bg-gradient-to-br from-slate-900 to-slate-800 text-slate-900 rounded-xl p-5 shadow-lg">
              <div className="flex items-center justify-between mb-5">
                <div>
                  <div className="text-xs uppercase tracking-wider text-blue-500 font-bold">Live Calculation</div>
                  <div className="text-sm text-blue-700 mt-0.5">Updates as you type</div>
                </div>
                <div className="w-8 h-8 bg-[#EFF6FF]0/20 rounded-lg flex items-center justify-center">
                  <Activity size={15} className="text-blue-500" />
                </div>
              </div>

              <div className="space-y-3 text-sm">
                <CalcRow label="Contract" value={fmt(c.contract, { currency: p.currency })} />
                <CalcRow label={`VAT (${fmtPct(p.vatPct)})`} value={`− ${fmt(c.vatAmount, { currency: p.currency })}`} tone="rose" />
                <CalcRow label="Net after VAT" value={fmt(c.netAfterVat, { currency: p.currency })} emphasis border />
                <CalcRow label={`Advance (${fmtPct(p.advancePct)})`} value={fmt(c.advanceAmount, { currency: p.currency })} />
                <CalcRow label="Bank retention" value={`− ${fmt(c.bankRetained, { currency: p.currency })}`} tone="rose" />
                <CalcRow label="Advance cash" value={fmt(c.advanceCashAvailable, { currency: p.currency })} tone="amber" emphasis border />
                <CalcRow label="Total deductions" value={`− ${fmt(c.totalDeductions, { currency: p.currency })}`} tone="rose" />
              </div>

              <div className="mt-5 pt-5 border-t border-blue-500/20 space-y-3">
                <div className="flex justify-between items-baseline">
                  <span className="text-xs uppercase tracking-wider text-[#2563EB] font-semibold">Real Cash</span>
                  <span className={`text-lg font-bold tabular-nums ${c.realCashAfterDeductions >= 0 ? 'text-blue-500' : 'text-rose-500'}`}>
                    {fmt(c.realCashAfterDeductions, { currency: p.currency })}
                  </span>
                </div>
                <div className="flex justify-between items-baseline">
                  <span className="text-xs uppercase tracking-wider text-[#2563EB] font-semibold">Final Profit</span>
                  <span className={`text-lg font-bold tabular-nums ${c.projectedFinalProfit >= 0 ? 'text-blue-500' : 'text-rose-500'}`}>
                    {fmt(c.projectedFinalProfit, { currency: p.currency })}
                  </span>
                </div>
              </div>

              <div className="mt-5 pt-5 border-t border-slate-700 flex items-center justify-between">
                <span className="text-xs uppercase tracking-wider text-[#2563EB] font-semibold">Assessment</span>
                <Pill tone={c.riskLevel}>
                  {c.riskLevel === 'safe' && <CheckCircle2 size={11} />}
                  {c.riskLevel === 'caution' && <AlertTriangle size={11} />}
                  {c.riskLevel === 'risky' && <XCircle size={11} />}
                  {c.riskLevel}
                </Pill>
              </div>
            </div>

            <Btn onClick={() => onSave(p)} icon={Save} fullWidth size="lg">Save Project</Btn>
          </div>
        </div>
      </div>
    </div>
  );
};

const CalcRow = ({ label, value, tone, emphasis, border }) => {
  const tones = { rose: 'text-rose-500', emerald: 'text-blue-500', amber: 'text-[#475569]' };
  return (
    <div className={`flex justify-between items-baseline ${border ? 'pb-3 border-b border-slate-700' : ''}`}>
      <span className={`text-xs ${emphasis ? 'text-slate-900 font-bold' : 'text-blue-700'}`}>{label}</span>
      <span className={`font-mono tabular-nums text-xs ${tones[tone] || 'text-slate-900'} ${emphasis ? 'font-bold' : ''}`}>{value}</span>
    </div>
  );
};

// ============================================================
// PROJECT DETAIL
// ============================================================
const ProjectDetail = ({ project, onBack, onEdit, onDelete, onDuplicate, onScenario }) => {
  const p = project;
  const c = calcProject(p);

  const deductionBreakdown = [
    { label: 'Mobilisation', v: +p.mobilisationCost || 0 },
    { label: 'Performance Bond', v: +p.performanceBondCost || 0 },
    { label: 'Bank Guarantee', v: +p.bankGuaranteeCost || 0 },
    { label: 'Insurance', v: +p.insuranceCost || 0 },
    { label: 'Tender / Application', v: +p.tenderCost || 0 },
    { label: 'Legal / Documentation', v: +p.legalCost || 0 },
    { label: 'Commissions', v: +p.commissionCost || 0 },
    ...(p.customDeductions || []).map(d => ({ label: d.label || 'Custom', v: +d.amount || 0 })),
    ...(p.partners || []).filter(pt => +pt.payout > 0).map(pt => ({ label: `Payout — ${pt.name || 'Partner'}`, v: +pt.payout || 0 }))
  ].filter(x => x.v > 0);

  const riskConfig = {
    safe: { bg: 'bg-[#EFF6FF]', border: 'border-[#CBD5E1]', text: 'text-blue-900', accent: 'text-blue-600', icon: CheckCircle2, title: 'Project is financially sound' },
    caution: { bg: 'bg-slate-50', border: 'border-slate-200', text: 'text-slate-900', accent: 'text-slate-600', icon: AlertTriangle, title: 'Proceed with caution' },
    risky: { bg: 'bg-rose-50', border: 'border-rose-200', text: 'text-rose-900', accent: 'text-rose-600', icon: XCircle, title: 'High risk — review carefully' }
  };
  const risk = riskConfig[c.riskLevel];

  return (
    <div>
      <div className="flex flex-wrap items-start justify-between gap-4 mb-6">
        <div className="flex items-start gap-3">
          <button onClick={onBack} className="p-2 hover:bg-[#F1F5F9] rounded-lg text-blue-700 transition-colors mt-0.5">
            <ArrowLeft size={18} />
          </button>
          <div>
            <div className="flex items-center gap-2 mb-2">
              <Pill tone={p.status}>{p.status}</Pill>
              <Pill tone={c.riskLevel}>
                {c.riskLevel === 'safe' && <CheckCircle2 size={11} />}
                {c.riskLevel === 'caution' && <AlertTriangle size={11} />}
                {c.riskLevel === 'risky' && <XCircle size={11} />}
                {c.riskLevel}
              </Pill>
            </div>
            <h1 className="font-display text-3xl font-extrabold text-slate-900">{p.name || 'Untitled Project'}</h1>
            <div className="flex flex-wrap items-center gap-x-4 gap-y-1 mt-2 text-sm text-[#2563EB]">
              {p.client && <span className="flex items-center gap-1.5"><Building2 size={13} />{p.client}</span>}
              {p.location && <span className="flex items-center gap-1.5"><MapPin size={13} />{p.location}</span>}
              {p.type && <span className="flex items-center gap-1.5"><FolderKanban size={13} />{p.type}</span>}
              {p.duration && <span className="flex items-center gap-1.5"><Clock size={13} />{p.duration}</span>}
            </div>
          </div>
        </div>
        <div className="flex gap-2 flex-wrap">
          <Btn onClick={() => openReportPDF(buildProjectReportHTML(p))} variant="secondary" size="sm" icon={FileText}>Export PDF</Btn>
          <Btn onClick={() => buildProjectXLSX(p)} variant="secondary" size="sm" icon={Download}>Export Excel</Btn>
          <Btn onClick={() => onScenario(p.id)} variant="secondary" size="sm" icon={Sliders}>Test Scenarios</Btn>
          <Btn onClick={() => onDuplicate(p.id)} variant="secondary" size="sm" icon={Copy}>Duplicate</Btn>
          <Btn onClick={() => onEdit(p.id)} variant="secondary" size="sm" icon={Edit2}>Edit</Btn>
          <Btn onClick={() => onDelete(p.id)} variant="danger" size="sm" icon={Trash2}>Delete</Btn>
        </div>
      </div>

      {/* BID DECISION VERDICT — auto-generated executive summary */}
      {(() => {
        const hs = computeHealthScore(p, c);
        const ds = computeDocStatus(p);
        let verdict, verdictBg, verdictBorder, verdictColor, verdictIcon, verdictTitle, verdictMsg;
        if (hs.score >= 70 && c.riskLevel !== 'risky' && c.realCashAfterDeductions > 0) {
          verdict = 'GO';
          verdictBg = 'from-emerald-500/10 via-emerald-500/5 to-transparent';
          verdictBorder = 'border-emerald-500/40';
          verdictColor = 'text-blue-600';
          verdictIcon = CheckCircle2;
          verdictTitle = 'Bid Recommended';
          verdictMsg = 'Strong financials, manageable risk, positive cash position.';
        } else if (hs.score >= 50 && c.realCashAfterDeductions > 0) {
          verdict = 'CAUTION';
          verdictBg = 'from-slate-500/10 via-slate-500/5 to-transparent';
          verdictBorder = 'border-slate-500/30';
          verdictColor = 'text-slate-400';
          verdictIcon = AlertTriangle;
          verdictTitle = 'Proceed With Caution';
          verdictMsg = 'Project is viable but needs review on at least one dimension before committing.';
        } else {
          verdict = 'REVIEW';
          verdictBg = 'from-rose-500/10 via-rose-500/5 to-transparent';
          verdictBorder = 'border-rose-500/30';
          verdictColor = 'text-rose-400';
          verdictIcon = XCircle;
          verdictTitle = 'Needs Significant Review';
          verdictMsg = 'Multiple financial or risk flags detected. Re-negotiate or restructure before signing.';
        }

        // Build supporting reasons
        const reasons = [];
        if (c.profitMargin >= 15) reasons.push({ icon: TrendingUp, color: 'text-emerald-600', text: `Strong ${c.profitMargin.toFixed(1)}% effective margin` });
        else if (c.profitMargin >= 8) reasons.push({ icon: TrendingUp, color: 'text-slate-400', text: `${c.profitMargin.toFixed(1)}% margin (acceptable)` });
        else reasons.push({ icon: TrendingDown, color: 'text-rose-400', text: `Thin ${c.profitMargin.toFixed(1)}% margin — verify` });

        if (c.realCashAfterDeductions > 0) reasons.push({ icon: DollarSign, color: 'text-emerald-600', text: `${fmtShort(c.realCashAfterDeductions, p.currency)} positive day-1 cash` });
        else reasons.push({ icon: AlertTriangle, color: 'text-rose-400', text: `Cash gap of ${fmtShort(Math.abs(c.realCashAfterDeductions), p.currency)}` });

        if (c.capitalEfficiency >= 2) reasons.push({ icon: Sparkles, color: 'text-emerald-600', text: `${c.capitalEfficiency.toFixed(2)}x capital efficiency` });
        else if (c.capitalEfficiency >= 1) reasons.push({ icon: Sparkles, color: 'text-slate-400', text: `${c.capitalEfficiency.toFixed(2)}x capital efficiency` });
        else if (c.capitalEfficiency > 0) reasons.push({ icon: Sparkles, color: 'text-rose-400', text: `Low ${c.capitalEfficiency.toFixed(2)}x capital efficiency` });

        if (ds.missing.length === 0 && ds.criticalTemplates.length > 0) reasons.push({ icon: FileCheck, color: 'text-emerald-600', text: 'All critical documents in place' });
        else if (ds.missing.length > 0) reasons.push({ icon: FileX, color: 'text-slate-400', text: `${ds.missing.length} critical document${ds.missing.length === 1 ? '' : 's'} missing` });

        const VerdictIcon = verdictIcon;

        return (
          <div className={`relative overflow-hidden mb-6 rounded-2xl border ${verdictBorder} bg-white bg-gradient-to-r ${verdictBg} backdrop-blur-sm`}>
            <div className="p-5 sm:p-6 grid grid-cols-1 md:grid-cols-3 gap-5 items-center">
              <div className="md:col-span-1 flex items-center gap-4">
                <div className={`w-16 h-16 rounded-2xl ${verdictBorder} border-2 bg-white flex items-center justify-center flex-shrink-0`}>
                  <VerdictIcon size={32} className={verdictColor} strokeWidth={2} />
                </div>
                <div>
                  <div className={`text-[10px] uppercase tracking-[0.18em] font-extrabold ${verdictColor} mb-1`}>Verdict</div>
                  <div className={`font-display text-3xl font-extrabold ${verdictColor} leading-none mb-1`}>{verdict}</div>
                  <div className="text-xs font-bold text-slate-900">{verdictTitle}</div>
                </div>
              </div>

              <div className="md:col-span-2 space-y-2.5">
                <div className="text-sm text-slate-700 font-bold leading-relaxed">{verdictMsg}</div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {reasons.slice(0, 4).map((r, i) => {
                    const RIcon = r.icon;
                    return (
                      <div key={i} className="flex items-center gap-2 bg-[#EFF6FF] border border-[#D8E4F8] rounded-lg px-2.5 py-1.5">
                        <RIcon size={13} className={r.color} strokeWidth={2.4} />
                        <span className="text-[11px] font-bold text-slate-800 leading-tight">{r.text}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Health score strip at bottom */}
            <div className="border-t border-[#D8E4F8] bg-[#F7F9FB] px-5 sm:px-6 py-3 flex items-center gap-4 flex-wrap">
              <div className="flex items-center gap-2">
                <Shield size={13} className="text-[#2563EB]" />
                <span className="text-[10px] uppercase tracking-wider font-bold text-[#2563EB]">Health Score</span>
              </div>
              <div className="flex-1 h-1.5 bg-[#EFF6FF] rounded-full overflow-hidden max-w-md">
                <div className={`h-full rounded-full transition-all ${hs.score >= 75 ? 'bg-[#EFF6FF]0' : hs.score >= 50 ? 'bg-slate-500' : 'bg-rose-500'}`} style={{ width: `${hs.score}%` }} />
              </div>
              <span className={`num-hero text-xl ${hs.score >= 75 ? 'text-emerald-600' : hs.score >= 50 ? 'text-slate-400' : 'text-rose-400'}`}>
                {hs.score}<span className="text-sm text-blue-600 ml-1">/ 100</span>
              </span>
            </div>
          </div>
        );
      })()}

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-4">
        <KpiTile help={GLOSSARY.contractValue} label="Contract Value" value={fmtShort(c.contract, p.currency)} icon={DollarSign} tone="emerald" />
        <KpiTile help={GLOSSARY.advanceCash} label="Advance Cash" value={fmtShort(c.advanceCashAvailable, p.currency)} sub={`After ${fmtPct(p.bankRetentionPct)} retention`} icon={Activity} tone="sky" />
        <KpiTile help={GLOSSARY.realCash} label="Real Cash" value={fmtShort(c.realCashAfterDeductions, p.currency)} sub="After deductions" icon={c.realCashAfterDeductions >= 0 ? TrendingUp : TrendingDown} tone={c.realCashAfterDeductions >= 0 ? 'teal' : 'rose'} />
        <KpiTile help={GLOSSARY.projectedProfit} label="Final Profit" value={fmtShort(c.projectedFinalProfit, p.currency)} sub="Projected net" icon={Trophy} tone={c.projectedFinalProfit >= 0 ? 'violet' : 'rose'} />
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <Card padding="p-4">
          <div className="flex items-center gap-2 mb-1.5">
            <Target size={13} className="text-[#2563EB]" />
            <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold flex items-center">
              Effective Margin <InfoIcon text={GLOSSARY.effectiveMargin} />
            </div>
          </div>
          <div className={`text-xl font-bold tabular-nums ${c.profitMargin >= 0 ? 'text-emerald-600' : 'text-rose-600'}`}>{c.profitMargin.toFixed(2)}%</div>
          <div className="text-[10px] text-blue-700 mt-0.5">Profit / Contract</div>
        </Card>
        <Card padding="p-4">
          <div className="flex items-center gap-2 mb-1.5">
            <Sparkles size={13} className="text-[#2563EB]" />
            <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold flex items-center">
              Capital Efficiency <InfoIcon text={GLOSSARY.capitalEfficiency} />
            </div>
          </div>
          <div className={`text-xl font-bold tabular-nums ${c.capitalEfficiency && c.capitalEfficiency >= 0 ? 'text-slate-900' : 'text-rose-600'}`}>
            {c.capitalEfficiency !== null ? `${c.capitalEfficiency.toFixed(2)}x` : '—'}
          </div>
          <div className="text-[10px] text-blue-700 mt-0.5">Profit / Deductions</div>
        </Card>
        <Card padding="p-4">
          <div className="flex items-center gap-2 mb-1.5">
            <DollarSign size={13} className="text-[#2563EB]" />
            <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold flex items-center">
              Expected Profit <InfoIcon text={GLOSSARY.expectedProfit} />
            </div>
          </div>
          <div className="text-xl font-bold text-slate-600 tabular-nums">{fmtShort(c.expectedProfit, p.currency)}</div>
          <div className="text-[10px] text-blue-700 mt-0.5">Before deductions</div>
        </Card>
        <Card padding="p-4">
          <div className="flex items-center gap-2 mb-1.5">
            <TrendingDown size={13} className="text-[#2563EB]" />
            <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold flex items-center">
              Total Deductions <InfoIcon text={GLOSSARY.totalDeductions} />
            </div>
          </div>
          <div className="text-xl font-bold text-rose-600 tabular-nums">{fmtShort(c.totalDeductions, p.currency)}</div>
          <div className="text-[10px] text-blue-700 mt-0.5">{((c.totalDeductions / c.contract) * 100).toFixed(1)}% of contract</div>
        </Card>
      </div>

      <div className={`${risk.bg} border ${risk.border} rounded-xl p-5 mb-6`}>
        <div className="flex items-start gap-3">
          <risk.icon className={`${risk.accent} flex-shrink-0 mt-0.5`} size={22} strokeWidth={2} />
          <div className="flex-1">
            <h3 className={`font-bold ${risk.text} mb-2`}>{risk.title}</h3>
            <ul className={`text-sm ${risk.text} space-y-1`}>
              {c.riskReasons.map((r, i) => (
                <li key={i} className="flex gap-2">
                  <span className={risk.accent}>•</span>
                  <span>{r}</span>
                </li>
              ))}
            </ul>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
        <Card>
          <h3 className="font-bold text-slate-900 mb-1">Calculation Walkthrough</h3>
          <p className="text-xs text-blue-700 mb-5">Step-by-step financial breakdown</p>
          <div className="space-y-3 text-sm">
            <WalkRow label="Contract value" value={fmt(c.contract, { currency: p.currency })} />
            <WalkRow label={`VAT (${fmtPct(p.vatPct)})`} value={`− ${fmt(c.vatAmount, { currency: p.currency })}`} tone="rose" />
            <WalkRow label="Net after VAT" value={fmt(c.netAfterVat, { currency: p.currency })} emphasis />
            <WalkRow label={`Expected profit (${fmtPct(p.profitPct)})`} value={fmt(c.expectedProfit, { currency: p.currency })} tone="amber" />
            <div className="border-t border-dashed border-[#D8E4F8] my-3" />
            <WalkRow label={`Advance (${fmtPct(p.advancePct)})`} value={fmt(c.advanceAmount, { currency: p.currency })} />
            <WalkRow label={`Bank retention (${fmtPct(p.bankRetentionPct)})`} value={`− ${fmt(c.bankRetained, { currency: p.currency })}`} tone="rose" />
            <WalkRow label="Advance cash available" value={fmt(c.advanceCashAvailable, { currency: p.currency })} emphasis tone="sky" />
            <div className="border-t border-dashed border-[#D8E4F8] my-3" />
            <WalkRow label="Total deductions" value={`− ${fmt(c.totalDeductions, { currency: p.currency })}`} tone="rose" />
            <div className="border-t-2 border-slate-900 my-3" />
            <WalkRow label="Real cash after deductions" value={fmt(c.realCashAfterDeductions, { currency: p.currency })} emphasis tone={c.realCashAfterDeductions >= 0 ? 'emerald' : 'rose'} large />
            <WalkRow label="Projected final profit" value={fmt(c.projectedFinalProfit, { currency: p.currency })} emphasis tone={c.projectedFinalProfit >= 0 ? 'emerald' : 'rose'} large />
          </div>
        </Card>

        <Card>
          <h3 className="font-bold text-slate-900 mb-1">Deductions Breakdown</h3>
          <p className="text-xs text-blue-700 mb-5">All costs & payouts</p>
          {deductionBreakdown.length === 0 ? (
            <div className="text-blue-700 italic text-sm text-center py-12 border-2 border-dashed border-[#D8E4F8] rounded-lg">
              No deductions entered
            </div>
          ) : (
            <div className="space-y-0">
              {deductionBreakdown.map((d, i) => (
                <div key={i} className="flex justify-between items-center text-sm py-3 border-b border-[#EEF2F5] last:border-0">
                  <span className="text-slate-900">{d.label}</span>
                  <span className="text-slate-900 font-bold tabular-nums">{fmt(d.v, { currency: p.currency })}</span>
                </div>
              ))}
              <div className="flex justify-between items-center pt-4 mt-2 border-t-2 border-slate-900">
                <span className="font-bold text-slate-900">Total Deductions</span>
                <span className="text-rose-600 font-bold tabular-nums">{fmt(c.totalDeductions, { currency: p.currency })}</span>
              </div>
            </div>
          )}
        </Card>
      </div>

      {c.partnersWithShare.length > 0 && (
        <Card className="mb-6">
          <h3 className="font-bold text-slate-900 mb-1">Partner Profit Share</h3>
          <p className="text-xs text-blue-700 mb-5">Distribution of final profit</p>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-[#D8E4F8]">
                  <th className="text-left py-3 text-xs uppercase tracking-wider text-[#2563EB] font-semibold">Partner</th>
                  <th className="text-right py-3 text-xs uppercase tracking-wider text-[#2563EB] font-semibold">Share %</th>
                  <th className="text-right py-3 text-xs uppercase tracking-wider text-[#2563EB] font-semibold">Direct Payout</th>
                  <th className="text-right py-3 text-xs uppercase tracking-wider text-[#2563EB] font-semibold">Profit Share</th>
                </tr>
              </thead>
              <tbody>
                {c.partnersWithShare.map((pt, i) => (
                  <tr key={i} className="border-b border-[#EEF2F5] last:border-0">
                    <td className="py-3 font-bold text-slate-900">{pt.name || 'Unnamed'}</td>
                    <td className="py-3 text-right text-slate-900 tabular-nums">{fmtPct(pt.sharePct)}</td>
                    <td className="py-3 text-right text-slate-900 tabular-nums">{fmt(pt.payout, { currency: p.currency })}</td>
                    <td className={`py-3 text-right font-bold tabular-nums ${pt.profitShare >= 0 ? 'text-blue-600' : 'text-rose-600'}`}>
                      {fmt(pt.profitShare, { currency: p.currency })}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {/* Health Score */}
      {(() => {
        const hs = computeHealthScore(p, c);
        const scoreColor = hs.score >= 75 ? 'emerald' : hs.score >= 50 ? 'amber' : 'rose';
        const colors = {
          emerald: { text: 'text-blue-600', bg: 'bg-[#EFF6FF]0', light: 'bg-[#EFF6FF]', border: 'border-[#CBD5E1]', label: 'Healthy' },
          amber: { text: 'text-slate-600', bg: 'bg-slate-500', light: 'bg-slate-50', border: 'border-slate-200', label: 'Needs attention' },
          rose: { text: 'text-rose-600', bg: 'bg-rose-500', light: 'bg-rose-50', border: 'border-rose-200', label: 'Critical' }
        };
        const c1 = colors[scoreColor];
        return (
          <Card className="mb-6">
            <div className="flex flex-wrap items-start gap-6">
              <div className="flex items-center gap-5">
                <div className="relative w-24 h-24 flex-shrink-0">
                  <svg className="w-24 h-24 -rotate-90">
                    <circle cx="48" cy="48" r="42" fill="none" stroke="#e2e8f0" strokeWidth="8" />
                    <circle cx="48" cy="48" r="42" fill="none" stroke="currentColor" strokeWidth="8" strokeLinecap="round" strokeDasharray={`${(hs.score / 100) * 264} 264`} className={c1.text} />
                  </svg>
                  <div className="absolute inset-0 flex items-center justify-center flex-col">
                    <div className={`text-2xl font-bold ${c1.text} tabular-nums`}>{hs.score}</div>
                    <div className="text-[9px] uppercase tracking-wider text-[#2563EB] font-semibold">/ 100</div>
                  </div>
                </div>
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <Target size={15} className="text-[#2563EB]" />
                    <h3 className="font-bold text-slate-900">Project Health Score</h3>
                  </div>
                  <div className={`text-sm font-bold ${c1.text} mb-1`}>{c1.label}</div>
                  <p className="text-xs text-blue-700 max-w-xs">Composite score from cash, profit, margin, capital efficiency, and coverage.</p>
                </div>
              </div>

              <div className="flex-1 min-w-[200px]">
                <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold mb-2">Score factors</div>
                <div className="space-y-1.5">
                  {hs.factors.map((f, i) => (
                    <div key={i} className="flex items-center justify-between text-xs">
                      <div className="flex items-center gap-2">
                        {f.good ? <CheckCircle2 size={12} className="text-blue-500" /> : <XCircle size={12} className="text-rose-500" />}
                        <span className="text-slate-900">{f.label}</span>
                      </div>
                      <span className={`font-mono font-bold tabular-nums ${f.good ? 'text-blue-600' : 'text-rose-600'}`}>
                        {f.delta > 0 ? '+' : ''}{f.delta}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </Card>
        );
      })()}

      {/* Stress Tests */}
      {c.contract > 0 && (
        <Card className="mb-6">
          <div className="flex items-start justify-between mb-4">
            <div>
              <div className="flex items-center gap-2 mb-1">
                <Sliders size={15} className="text-[#2563EB]" />
                <h3 className="font-bold text-slate-900">Stress Tests</h3>
              </div>
              <p className="text-xs text-[#2563EB]">How the project holds up under adverse conditions</p>
            </div>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {computeStressTests(p).map((t, i) => {
              const profitColor = t.profit >= 0 ? 'text-blue-600' : 'text-rose-600';
              const cashColor = t.cash >= 0 ? 'text-blue-600' : 'text-rose-600';
              return (
                <div key={i} className="bg-[#EFF6FF] border border-[#D8E4F8] rounded-lg p-4">
                  <div className="flex items-start justify-between mb-3">
                    <div className="min-w-0 flex-1">
                      <div className="font-bold text-sm text-slate-900">{t.label}</div>
                      <div className="text-[10px] text-blue-700 mt-0.5">{t.sub}</div>
                    </div>
                    <Pill tone={t.risk} size="sm">{t.risk}</Pill>
                  </div>
                  <div className="grid grid-cols-2 gap-3 text-xs">
                    <div>
                      <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold mb-0.5">Profit</div>
                      <div className={`font-bold tabular-nums ${profitColor}`}>{fmtShort(t.profit, p.currency)}</div>
                      <div className="text-[10px] text-blue-700 tabular-nums mt-0.5">
                        {t.profitDelta < 0 ? '↓' : '↑'} {fmtShort(Math.abs(t.profitDelta), p.currency)}
                      </div>
                    </div>
                    <div>
                      <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold mb-0.5">Real Cash</div>
                      <div className={`font-bold tabular-nums ${cashColor}`}>{fmtShort(t.cash, p.currency)}</div>
                      <div className="text-[10px] text-blue-700 tabular-nums mt-0.5">
                        {t.cashDelta < 0 ? '↓' : '↑'} {fmtShort(Math.abs(t.cashDelta), p.currency)}
                      </div>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
      )}

      {/* Documents */}
      {(() => {
        const ds = computeDocStatus(p);
        const docs = p.documents || [];
        if (docs.length === 0 && ds.criticalTemplates.length === 0) return null;
        return (
          <Card className="mb-6">
            <div className="flex items-start justify-between mb-4 flex-wrap gap-3">
              <div>
                <div className="flex items-center gap-2 mb-1">
                  <FileCheck size={15} className="text-blue-500" />
                  <h3 className="font-bold text-slate-900">Documents</h3>
                </div>
                <p className="text-xs text-[#2563EB]">Compliance status and document tracking</p>
              </div>
              <div className="flex items-center gap-3">
                <div className="text-right">
                  <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold">Compliance</div>
                  <div className={`text-xl font-bold tabular-nums ${ds.compliancePct >= 80 ? 'text-blue-600' : ds.compliancePct >= 50 ? 'text-slate-600' : 'text-rose-600'}`}>
                    {ds.compliancePct.toFixed(0)}%
                  </div>
                </div>
                <Btn onClick={() => onEdit(p.id)} variant="secondary" size="sm" icon={Edit2}>Manage</Btn>
              </div>
            </div>

            <div className="h-2 bg-[#EFF6FF] rounded-full overflow-hidden mb-4">
              <div className={`h-full rounded-full transition-all ${ds.compliancePct >= 80 ? 'bg-[#EFF6FF]0' : ds.compliancePct >= 50 ? 'bg-slate-500' : 'bg-rose-500'}`} style={{ width: `${ds.compliancePct}%` }} />
            </div>

            {/* Status grid */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-5">
              <div className="bg-[#EFF6FF] border border-[#CBD5E1] rounded-lg p-3">
                <div className="flex items-center gap-1.5 mb-1">
                  <FileCheck size={13} className="text-blue-500" />
                  <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold">Received</div>
                </div>
                <div className="text-lg font-bold text-blue-900 tabular-nums">{docs.filter(d => d.status === 'received').length}</div>
              </div>
              <div className="bg-slate-50 border border-slate-200 rounded-lg p-3">
                <div className="flex items-center gap-1.5 mb-1">
                  <FileClock size={13} className="text-slate-600" />
                  <div className="text-[10px] uppercase tracking-wider text-slate-700 font-bold">Pending</div>
                </div>
                <div className="text-lg font-bold text-slate-900 tabular-nums">{docs.filter(d => d.status === 'pending' || d.status === 'required').length + (ds.criticalTemplates.length - ds.criticalReceived.length - docs.filter(d => (d.status === 'pending' || d.status === 'required') && d.templateId && DOCUMENT_TEMPLATES.find(t => t.id === d.templateId)?.critical).length)}</div>
              </div>
              <div className="bg-rose-50 border border-rose-200 rounded-lg p-3">
                <div className="flex items-center gap-1.5 mb-1">
                  <FileX size={13} className="text-rose-600" />
                  <div className="text-[10px] uppercase tracking-wider text-rose-700 font-bold">Expired</div>
                </div>
                <div className="text-lg font-bold text-rose-900 tabular-nums">{ds.expired.length}</div>
              </div>
              <div className="bg-[#EFF6FF] border border-[#D8E4F8] rounded-lg p-3">
                <div className="flex items-center gap-1.5 mb-1">
                  <Calendar size={13} className="text-[#2563EB]" />
                  <div className="text-[10px] uppercase tracking-wider text-slate-900 font-bold">Expiring</div>
                </div>
                <div className="text-lg font-bold text-slate-900 tabular-nums">{ds.expiringSoon.length}</div>
              </div>
            </div>

            {/* Missing critical docs alert */}
            {ds.missing.length > 0 && (
              <div className="bg-rose-50 border border-rose-200 rounded-lg p-3 mb-4">
                <div className="flex items-start gap-2">
                  <AlertTriangle size={15} className="text-rose-600 flex-shrink-0 mt-0.5" />
                  <div>
                    <div className="text-sm font-bold text-rose-900 mb-1">{ds.missing.length} critical document{ds.missing.length === 1 ? '' : 's'} missing</div>
                    <div className="text-xs text-rose-800 leading-relaxed">{ds.missing.map(t => t.label).join(', ')}</div>
                  </div>
                </div>
              </div>
            )}

            {/* Quick links to received docs */}
            {docs.filter(d => d.status === 'received' && (d.url || d.fileName)).length > 0 && (
              <div>
                <div className="text-[11px] uppercase tracking-wider text-[#2563EB] font-semibold mb-2">Quick Access</div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {docs.filter(d => d.status === 'received' && (d.url || d.fileName)).slice(0, 12).map((d, i) => {
                    const t = d.templateId ? getDocTemplate(d.templateId) : null;
                    const label = t?.label || d.customLabel || 'Document';
                    const docKey = d.templateId || d.id;
                    if (d.fileName) {
                      return (
                        <div key={i} className="flex items-center gap-2 px-3 py-2 bg-white border border-[#CBD5E1] rounded-lg hover:border-[#2563EB]/50 hover:bg-[#F1F5F9]/40 transition-all group">
                          <File size={13} className="text-blue-600 flex-shrink-0" />
                          <div className="flex-1 min-w-0">
                            <div className="text-xs font-bold text-slate-900 truncate group-hover:text-[#2563EB]">{label}</div>
                            <div className="text-[10px] text-blue-700 truncate">{d.fileName} · {formatBytes(d.fileSize)}</div>
                          </div>
                          <button onClick={() => previewDocFile(p.id, docKey)} title="Preview" className="p-1 text-blue-700 hover:text-[#2563EB] rounded"><Eye size={12} /></button>
                          <button onClick={() => downloadDocFile(p.id, docKey, d.fileName)} title="Download" className="p-1 text-blue-700 hover:text-[#2563EB] rounded"><Download size={12} /></button>
                        </div>
                      );
                    }
                    return (
                      <a key={i} href={d.url} target="_blank" rel="noopener noreferrer" className="flex items-center gap-2 px-3 py-2 bg-white border border-[#CBD5E1] rounded-lg hover:border-[#2563EB]/50 hover:bg-[#F1F5F9]/40 transition-all group">
                        <ExternalLink size={13} className="text-blue-600 flex-shrink-0" />
                        <span className="text-xs font-bold text-slate-900 truncate group-hover:text-[#2563EB]">{label}</span>
                      </a>
                    );
                  })}
                </div>
              </div>
            )}
          </Card>
        );
      })()}

      {/* Cash Flow Timeline */}
      {(() => {
        const cf = buildCashflow(p, c);
        const chartData = cf.periods.map(x => ({
          label: x.label,
          inflow: Math.round(x.inflow / 1_000_000),
          outflow: -Math.round(x.outflow / 1_000_000),
          cumulative: Math.round(x.cumulative / 1_000_000)
        }));
        return (
          <Card className="mb-6">
            <div className="flex items-start justify-between mb-5 flex-wrap gap-3">
              <div>
                <div className="flex items-center gap-2 mb-1">
                  <Activity size={15} className="text-blue-500" />
                  <h3 className="font-display font-bold text-slate-900">Cash Flow Projection</h3>
                </div>
                <p className="text-xs text-[#2563EB]">Month-by-month inflows, outflows, and cumulative position over {cf.months} months</p>
              </div>
              <div className="flex gap-3 text-xs">
                <div className="text-right">
                  <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold">Min Cash</div>
                  <div className={`num-hero text-lg ${cf.minCash < 0 ? 'text-rose-500' : 'text-blue-500'}`}>{fmtShort(cf.minCash * 1_000_000)}</div>
                </div>
                {cf.breakEvenMonth !== undefined && (
                  <div className="text-right">
                    <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold">Break-even</div>
                    <div className="num-hero text-lg text-slate-900">M{cf.breakEvenMonth}</div>
                  </div>
                )}
              </div>
            </div>

            {cf.minCash < 0 && (
              <div className="bg-[#FFE4E6] border border-[#FECDD3] rounded-lg p-3 mb-4 flex items-start gap-2">
                <AlertTriangle size={14} className="text-rose-500 flex-shrink-0 mt-0.5" />
                <div className="text-xs text-[#7A1F32]">
                  <strong className="text-[#7A1F32]">Cash gap warning:</strong> Cumulative cash dips to {fmtShort(cf.minCash)} at M{cf.peakNegative.month}. You will need bridge financing or a higher advance to cover this period.
                </div>
              </div>
            )}

            <div className="bg-[#EFF6FF] border border-[#EEF2F5] rounded-lg p-3" style={{ height: 280 }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={chartData} margin={{ top: 10, right: 10, left: -10, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#D8E4F8" />
                  <XAxis dataKey="label" tick={{ fill: '#697386', fontSize: 10 }} stroke="#CBD5E1" />
                  <YAxis tick={{ fill: '#697386', fontSize: 10 }} stroke="#CBD5E1" tickFormatter={(v) => `${v}M`} />
                  <RechartsTooltip
                    contentStyle={{ background: '#FFFFFF', border: '1px solid #D8E4F8', color: '#1A1F36', borderRadius: 8, fontSize: 11 }}
                    labelStyle={{ color: '#1A1F36', fontWeight: 700 }}
                    formatter={(v, name) => [`${p.currency || 'KES'} ${Math.abs(v).toLocaleString()}M`, name]}
                  />
                  <Bar dataKey="inflow" fill="#2563EB" name="Inflow" radius={[2,2,0,0]} />
                  <Bar dataKey="outflow" fill="#F43F5E" name="Outflow" radius={[2,2,0,0]} />
                  <Bar dataKey="cumulative" fill="#2563EB" name="Cumulative" radius={[2,2,0,0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-4 pt-4 border-t border-[#D8E4F8]">
              <div>
                <div className="text-[9px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1">Monthly Billing</div>
                <div className="num-hero text-sm text-blue-500">{fmtShort(cf.monthlyBilling, p.currency)}</div>
              </div>
              <div>
                <div className="text-[9px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1">Monthly Op Cost</div>
                <div className="num-hero text-sm text-[#475569]">{fmtShort(cf.monthlyOpCost, p.currency)}</div>
              </div>
              <div>
                <div className="text-[9px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1">Final Cash Position</div>
                <div className={`num-hero text-sm ${cf.periods[cf.periods.length-1].cumulative >= 0 ? 'text-blue-500' : 'text-rose-500'}`}>{fmtShort(cf.periods[cf.periods.length-1].cumulative, p.currency)}</div>
              </div>
              <div>
                <div className="text-[9px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1">Duration</div>
                <div className="num-hero text-sm text-slate-900">{cf.months} mo</div>
              </div>
            </div>
          </Card>
        );
      })()}

      {p.notes && (
        <Card>
          <div className="flex items-center gap-2 mb-3">
            <StickyNote size={15} className="text-[#2563EB]" />
            <h3 className="font-bold text-slate-900">Notes</h3>
          </div>
          <div className="text-sm text-slate-900 whitespace-pre-wrap leading-relaxed bg-slate-50 border-l-4 border-slate-400 p-4 rounded-r-lg">{p.notes}</div>
        </Card>
      )}
    </div>
  );
};

const WalkRow = ({ label, value, tone, emphasis, large }) => {
  const tones = { rose: 'text-rose-600', emerald: 'text-blue-600', amber: 'text-slate-600', sky: 'text-blue-700' };
  return (
    <div className="flex justify-between items-baseline">
      <span className={`${emphasis ? 'text-slate-900 font-bold' : 'text-blue-700'} text-sm`}>{label}</span>
      <span className={`font-mono tabular-nums ${tones[tone] || 'text-slate-900'} ${emphasis ? 'font-bold' : ''} ${large ? 'text-base' : 'text-sm'}`}>{value}</span>
    </div>
  );
};

// ============================================================
// COMPARE
// ============================================================
const Compare = ({ projects }) => {
  const [selected, setSelected] = useState([]);
  const toggle = id => setSelected(selected.includes(id) ? selected.filter(x => x !== id) : selected.length < 4 ? [...selected, id] : selected);
  const compared = projects.filter(p => selected.includes(p.id)).map(p => ({ p, c: calcProject(p) }));

  const metrics = [
    { key: 'contract', label: 'Contract Value', higher: true, get: x => x.c.contract },
    { key: 'advanceCashAvailable', label: 'Upfront Cash', higher: true, get: x => x.c.advanceCashAvailable },
    { key: 'totalDeductions', label: 'Total Deductions', higher: false, get: x => x.c.totalDeductions },
    { key: 'realCashAfterDeductions', label: 'Real Cash After Deductions', higher: true, get: x => x.c.realCashAfterDeductions },
    { key: 'projectedFinalProfit', label: 'Final Profit', higher: true, get: x => x.c.projectedFinalProfit },
    { key: 'profitMargin', label: 'Effective Margin', higher: true, get: x => x.c.contract > 0 ? (x.c.projectedFinalProfit / x.c.contract) * 100 : 0, fmt: v => `${v.toFixed(1)}%` }
  ];

  const bestOverall = compared.length > 0 ? [...compared].sort((a, b) => b.c.projectedFinalProfit - a.c.projectedFinalProfit)[0] : null;

  return (
    <div>
      <PageHeader title="Compare Projects" subtitle="Select up to 4 projects to compare side by side" />

      <Card className="mb-6" padding="p-5">
        <h3 className="font-bold text-slate-900 mb-4">Select Projects to Compare</h3>
        {projects.length === 0 ? (
          <div className="text-sm text-blue-700 italic">No projects available. Create some projects first.</div>
        ) : (
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-3">
            {projects.map(p => {
              const on = selected.includes(p.id);
              return (
                <button
                  key={p.id}
                  onClick={() => toggle(p.id)}
                  className={`text-left p-3 border-2 rounded-lg transition-all ${
                    on
                      ? 'bg-[#EFF6FF] border-blue-500'
                      : 'bg-white border-[#D8E4F8] hover:border-slate-400'
                  }`}
                >
                  <div className={`font-bold text-sm truncate ${on ? 'text-blue-900' : 'text-slate-900'}`}>{p.name || 'Untitled'}</div>
                  <div className="text-xs text-blue-700 truncate mt-0.5">{p.client || 'No client'}</div>
                </button>
              );
            })}
          </div>
        )}
      </Card>

      {compared.length < 2 ? (
        <EmptyState
          icon={GitCompare}
          title="Select projects to compare"
          subtitle="Choose at least 2 projects above to see a side-by-side analysis with winners highlighted."
        />
      ) : (
        <>
          <Card className="mb-6" padding="p-0">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-[#D8E4F8] bg-[#EFF6FF]">
                    <th className="text-left px-5 py-4 text-xs uppercase tracking-wider text-[#2563EB] font-semibold">Metric</th>
                    {compared.map(({ p }) => (
                      <th key={p.id} className="text-right px-5 py-4 font-bold text-slate-900">{p.name || 'Untitled'}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {metrics.map(m => {
                    const values = compared.map(x => m.get(x));
                    const best = m.higher ? Math.max(...values) : Math.min(...values);
                    return (
                      <tr key={m.key} className="border-b border-[#EEF2F5] last:border-0">
                        <td className="px-5 py-4 text-slate-900">{m.label}</td>
                        {compared.map(({ p }, i) => {
                          const isBest = values[i] === best;
                          return (
                            <td key={p.id} className={`px-5 py-4 text-right font-mono tabular-nums ${isBest ? 'bg-[#EFF6FF] text-[#2563EB] font-semibold' : 'text-slate-900'}`}>
                              {m.fmt ? m.fmt(values[i]) : fmtShort(values[i], p.currency)}
                              {isBest && <Trophy size={12} className="inline ml-1.5 text-[#2563EB]" />}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                  <tr>
                    <td className="px-5 py-4 text-slate-900">Risk Level</td>
                    {compared.map(({ p, c }) => (
                      <td key={p.id} className="px-5 py-4 text-right"><Pill tone={c.riskLevel}>{c.riskLevel}</Pill></td>
                    ))}
                  </tr>
                </tbody>
              </table>
            </div>
          </Card>

          {bestOverall && (
            <div className="bg-gradient-to-br from-blue-600 to-blue-700 text-slate-900 rounded-xl p-6 shadow-lg">
              <div className="flex items-start gap-4">
                <div className="w-12 h-12 bg-white/20 rounded-xl flex items-center justify-center flex-shrink-0">
                  <Trophy size={22} className="text-slate-900" />
                </div>
                <div>
                  <div className="text-xs uppercase tracking-wider text-blue-100 font-bold mb-1">Recommendation</div>
                  <div className="text-lg font-bold mb-2">{bestOverall.p.name} is the strongest option</div>
                  <p className="text-sm text-blue-50 leading-relaxed">
                    Projected final profit of <span className="font-bold text-slate-900">{fmt(bestOverall.c.projectedFinalProfit, { currency: bestOverall.p.currency })}</span>. Cross-reference against upfront cash needs and risk level before committing capital.
                  </p>
                </div>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
};

// ============================================================
// REPORTS
// ============================================================
const Reports = ({ projects }) => {
  const [type, setType] = useState('summary');
  const [selected, setSelected] = useState([]);
  const data = useMemo(() => projects.map(p => ({ p, c: calcProject(p) })), [projects]);

  const toggleProject = (id) => setSelected(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
  const selectAll = () => setSelected(projects.map(p => p.id));
  const clearAll = () => setSelected([]);

  const customSelected = projects.filter(p => selected.includes(p.id));

  const exportCSV = () => {
    let csv = '';
    if (type === 'summary') {
      csv = 'Project,Client,Location,Status,Contract,Cash,Profit,Risk\n';
      data.forEach(({ p, c }) => { csv += `"${p.name}","${p.client}","${p.location}",${p.status},${c.contract},${c.realCashAfterDeductions},${c.projectedFinalProfit},${c.riskLevel}\n`; });
    } else if (type === 'deductions') {
      csv = 'Project,Mobilisation,PerformanceBond,BankGuarantee,Insurance,Tender,Legal,Commission,Custom,Partners,Total\n';
      data.forEach(({ p, c }) => { csv += `"${p.name}",${p.mobilisationCost||0},${p.performanceBondCost||0},${p.bankGuaranteeCost||0},${p.insuranceCost||0},${p.tenderCost||0},${p.legalCost||0},${p.commissionCost||0},${c.customTotal},${c.partnerTotal},${c.totalDeductions}\n`; });
    } else if (type === 'partners') {
      csv = 'Project,Partner,SharePct,DirectPayout,ProfitShare\n';
      data.forEach(({ p, c }) => { c.partnersWithShare.forEach(pt => { csv += `"${p.name}","${pt.name}",${pt.sharePct}%,${pt.payout},${pt.profitShare}\n`; }); });
    } else if (type === 'cashflow') {
      csv = 'Project,Contract,Advance,BankRetention,AdvanceCash,Deductions,RealCash,FinalProfit\n';
      data.forEach(({ p, c }) => { csv += `"${p.name}",${c.contract},${c.advanceAmount},${c.bankRetained},${c.advanceCashAvailable},${c.totalDeductions},${c.realCashAfterDeductions},${c.projectedFinalProfit}\n`; });
    } else if (type === 'profit') {
      csv = 'Project,Contract,ExpectedProfit,Deductions,FinalProfit,MarginPct\n';
      data.forEach(({ p, c }) => { const margin = c.contract > 0 ? (c.projectedFinalProfit / c.contract * 100).toFixed(2) : 0; csv += `"${p.name}",${c.contract},${c.expectedProfit},${c.totalDeductions},${c.projectedFinalProfit},${margin}%\n`; });
    }
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `rays-report-${type}-${Date.now()}.csv`;
    a.click();
  };

  const types = [
    { v: 'summary', l: 'Project Summary' },
    { v: 'profit', l: 'Profit Projection' },
    { v: 'deductions', l: 'Deduction Breakdown' },
    { v: 'partners', l: 'Partner Profit' },
    { v: 'cashflow', l: 'Cashflow' }
  ];

  return (
    <div>
      <PageHeader
        title="Reports"
        subtitle="Generate executive-level financial reports in PDF or Excel"
      />

      {projects.length === 0 ? (
        <EmptyState icon={FileText} title="No projects to report on" subtitle="Create some projects first to generate reports." />
      ) : (
        <>
          {/* Executive download center */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-6">
            {/* Single project */}
            <Card className="hover:shadow-md transition-shadow">
              <div className="flex items-center gap-2 mb-3">
                <div className="w-9 h-9 bg-[#DBEAFE] rounded-lg flex items-center justify-center">
                  <FileText size={17} className="text-[#2563EB]" />
                </div>
                <div>
                  <div className="font-display font-extrabold text-slate-900 text-base">Per-Project Report</div>
                  <div className="text-[11px] text-[#2563EB]">Full breakdown of one project</div>
                </div>
              </div>
              <p className="text-xs text-blue-700 mb-3 leading-relaxed">Executive summary, financial walkthrough, deductions, partners, documents — for a single project.</p>
              <div className="text-[10px] text-blue-700 mb-2 font-bold uppercase tracking-wider">Open project then tap "Export"</div>
            </Card>

            {/* All projects */}
            <Card className="hover:shadow-md transition-shadow border-[#CBD5E1] bg-gradient-to-br from-blue-50/40 to-white">
              <div className="flex items-center gap-2 mb-3">
                <div className="w-9 h-9 bg-blue-600 rounded-lg flex items-center justify-center">
                  <Database size={17} className="text-slate-900" />
                </div>
                <div>
                  <div className="font-display font-extrabold text-slate-900 text-base">Portfolio Report</div>
                  <div className="text-[11px] text-[#2563EB] font-semibold">All {projects.length} projects</div>
                </div>
              </div>
              <p className="text-xs text-blue-700 mb-3 leading-relaxed">Combined portfolio view with KPIs, ledger, risk distribution, and recommendations.</p>
              <div className="flex gap-2">
                <Btn onClick={() => openReportPDF(buildPortfolioReportHTML(projects, 'Full Portfolio Report'))} size="sm" icon={FileText} fullWidth>PDF</Btn>
                <Btn onClick={() => buildPortfolioXLSX(projects, 'Full Portfolio')} variant="secondary" size="sm" icon={Download} fullWidth>Excel</Btn>
              </div>
            </Card>

            {/* Custom selection */}
            <Card className="hover:shadow-md transition-shadow">
              <div className="flex items-center gap-2 mb-3">
                <div className="w-9 h-9 bg-slate-100 rounded-lg flex items-center justify-center">
                  <GitCompare size={17} className="text-slate-700" />
                </div>
                <div>
                  <div className="font-display font-extrabold text-slate-900 text-base">Custom Report</div>
                  <div className="text-[11px] text-[#2563EB]">{selected.length === 0 ? 'Pick projects below' : `${selected.length} selected`}</div>
                </div>
              </div>
              <p className="text-xs text-blue-700 mb-3 leading-relaxed">Mix and match selected projects into a combined report — useful for investor decks or partner reviews.</p>
              <div className="flex gap-2">
                <Btn onClick={() => openReportPDF(buildPortfolioReportHTML(customSelected, `Custom Report (${customSelected.length} projects)`))} disabled={customSelected.length === 0} size="sm" icon={FileText} fullWidth>PDF</Btn>
                <Btn onClick={() => buildPortfolioXLSX(customSelected, 'Custom Report')} disabled={customSelected.length === 0} variant="secondary" size="sm" icon={Download} fullWidth>Excel</Btn>
              </div>
            </Card>
          </div>

          {/* Custom project selector */}
          <Card className="mb-6">
            <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
              <div>
                <div className="font-display font-bold text-slate-900">Select Projects for Custom Report</div>
                <div className="text-xs text-[#2563EB]">Tap any project to include or exclude</div>
              </div>
              <div className="flex gap-2">
                <button onClick={selectAll} className="text-xs font-bold text-blue-700 hover:text-blue-900 px-3 py-1.5 border border-[#CBD5E1] rounded bg-[#EFF6FF]">Select all</button>
                <button onClick={clearAll} className="text-xs font-bold text-blue-700 hover:text-slate-900 px-3 py-1.5 border border-[#D8E4F8] rounded bg-white">Clear</button>
              </div>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2">
              {projects.map(p => {
                const on = selected.includes(p.id);
                return (
                  <button key={p.id} onClick={() => toggleProject(p.id)} className={`text-left p-3 border-2 rounded-lg transition-all ${on ? 'bg-[#EFF6FF] border-slate-500' : 'bg-white border-[#D8E4F8] hover:border-slate-400'}`}>
                    <div className={`font-bold text-sm truncate ${on ? 'text-slate-900' : 'text-slate-900'}`}>{p.name || 'Untitled'}</div>
                    <div className="text-xs text-blue-700 truncate mt-0.5">{p.client || 'No client'}</div>
                  </button>
                );
              })}
            </div>
          </Card>

          {/* Quick CSV export */}
          <div className="mb-6">
            <h3 className="font-display font-bold text-slate-900 mb-1">Quick Data Tables</h3>
            <p className="text-xs text-blue-700 mb-4">View raw tables and export as CSV for spreadsheet analysis</p>
            <div className="flex flex-wrap gap-2 mb-4">
              {types.map(t => (
                <button
                  key={t.v}
                  onClick={() => setType(t.v)}
                  className={`px-4 py-2 text-sm font-bold rounded-lg transition-all ${
                    type === t.v
                      ? 'bg-blue-600 text-slate-900 shadow-sm'
                      : 'bg-white border border-[#CBD5E1] text-slate-900 hover:border-slate-300'
                  }`}
                >
                  {t.l}
                </button>
              ))}
              <Btn onClick={exportCSV} icon={Download} size="md" variant="secondary">Export CSV</Btn>
            </div>

            <Card padding="p-0">
              <div className="overflow-x-auto">
                {type === 'summary' && <SummaryTable data={data} />}
                {type === 'profit' && <ProfitTable data={data} />}
                {type === 'deductions' && <DeductionsTable data={data} />}
                {type === 'partners' && <PartnersTable data={data} />}
                {type === 'cashflow' && <CashflowTable data={data} />}
              </div>
            </Card>
          </div>
        </>
      )}
    </div>
  );
};

const TH = ({ children, align = 'left', help }) => (
  <th className={`px-5 py-4 text-xs uppercase tracking-wider text-[#2563EB] font-semibold text-${align}`}>
    <span className="inline-flex items-center">{children}{help && <InfoIcon text={help} />}</span>
  </th>
);
const TD = ({ children, align = 'left', mono, tone, bold }) => {
  const tones = { rose: 'text-rose-600', emerald: 'text-blue-600', amber: 'text-slate-600' };
  return <td className={`px-5 py-4 text-${align} ${mono ? 'font-mono text-sm tabular-nums' : 'text-sm'} ${tones[tone] || 'text-slate-900'} ${bold ? 'font-bold' : ''}`}>{children}</td>;
};

const SummaryTable = ({ data }) => (
  <table className="w-full">
    <thead><tr className="border-b border-[#D8E4F8] bg-[#EFF6FF]"><TH>Project</TH><TH>Client</TH><TH help={GLOSSARY.planning}>Status</TH><TH align="right" help={GLOSSARY.contractValue}>Contract</TH><TH align="right" help={GLOSSARY.projectedProfit}>Profit</TH><TH align="center" help={GLOSSARY.riskLevel}>Risk</TH></tr></thead>
    <tbody>{data.map(({ p, c }) => (
      <tr key={p.id} className="border-b border-[#EEF2F5] last:border-0 hover:bg-[#F1F5F9]/50">
        <TD bold>{p.name || 'Untitled'}</TD>
        <TD>{p.client || '—'}</TD><TD><Pill tone={p.status} size="sm">{p.status}</Pill></TD>
        <TD align="right" mono>{fmt(c.contract, { currency: p.currency })}</TD>
        <TD align="right" mono bold tone={c.projectedFinalProfit >= 0 ? 'emerald' : 'rose'}>{fmt(c.projectedFinalProfit, { currency: p.currency })}</TD>
        <TD align="center"><Pill tone={c.riskLevel} size="sm">{c.riskLevel}</Pill></TD>
      </tr>
    ))}</tbody>
  </table>
);

const ProfitTable = ({ data }) => (
  <table className="w-full">
    <thead><tr className="border-b border-[#D8E4F8] bg-[#EFF6FF]"><TH>Project</TH><TH align="right" help={GLOSSARY.contractValue}>Contract</TH><TH align="right" help={GLOSSARY.expectedProfit}>Expected</TH><TH align="right" help={GLOSSARY.totalDeductions}>Deductions</TH><TH align="right" help={GLOSSARY.projectedProfit}>Final Profit</TH><TH align="right" help={GLOSSARY.effectiveMargin}>Margin</TH></tr></thead>
    <tbody>{data.map(({ p, c }) => {
      const margin = c.contract > 0 ? (c.projectedFinalProfit / c.contract * 100) : 0;
      return (
        <tr key={p.id} className="border-b border-[#EEF2F5] last:border-0 hover:bg-[#F1F5F9]/50">
          <TD bold>{p.name || 'Untitled'}</TD>
          <TD align="right" mono>{fmt(c.contract, { currency: p.currency })}</TD>
          <TD align="right" mono tone="amber">{fmt(c.expectedProfit, { currency: p.currency })}</TD>
          <TD align="right" mono tone="rose">{fmt(c.totalDeductions, { currency: p.currency })}</TD>
          <TD align="right" mono bold tone={c.projectedFinalProfit >= 0 ? 'emerald' : 'rose'}>{fmt(c.projectedFinalProfit, { currency: p.currency })}</TD>
          <TD align="right" mono bold tone={margin >= 0 ? 'emerald' : 'rose'}>{margin.toFixed(1)}%</TD>
        </tr>
      );
    })}</tbody>
  </table>
);

const DeductionsTable = ({ data }) => (
  <table className="w-full">
    <thead><tr className="border-b border-[#D8E4F8] bg-[#EFF6FF]"><TH>Project</TH><TH align="right">Mobilisation</TH><TH align="right">Bond</TH><TH align="right">Guarantee</TH><TH align="right">Insurance</TH><TH align="right">Tender</TH><TH align="right">Legal</TH><TH align="right">Commissions</TH><TH align="right">Total</TH></tr></thead>
    <tbody>{data.map(({ p, c }) => (
      <tr key={p.id} className="border-b border-[#EEF2F5] last:border-0 hover:bg-[#F1F5F9]/50">
        <TD bold>{p.name || 'Untitled'}</TD>
        <TD align="right" mono>{fmtShort(p.mobilisationCost, p.currency)}</TD>
        <TD align="right" mono>{fmtShort(p.performanceBondCost, p.currency)}</TD>
        <TD align="right" mono>{fmtShort(p.bankGuaranteeCost, p.currency)}</TD>
        <TD align="right" mono>{fmtShort(p.insuranceCost, p.currency)}</TD>
        <TD align="right" mono>{fmtShort(p.tenderCost, p.currency)}</TD>
        <TD align="right" mono>{fmtShort(p.legalCost, p.currency)}</TD>
        <TD align="right" mono>{fmtShort(p.commissionCost, p.currency)}</TD>
        <TD align="right" mono bold tone="rose">{fmtShort(c.totalDeductions, p.currency)}</TD>
      </tr>
    ))}</tbody>
  </table>
);

const PartnersTable = ({ data }) => {
  const rows = data.flatMap(({ p, c }) => c.partnersWithShare.map(pt => ({ p, pt })));
  if (rows.length === 0) return <div className="p-12 text-center text-blue-700 italic">No partners across any project</div>;
  return (
    <table className="w-full">
      <thead><tr className="border-b border-[#D8E4F8] bg-[#EFF6FF]"><TH>Project</TH><TH>Partner</TH><TH align="right">Share %</TH><TH align="right">Direct Payout</TH><TH align="right">Profit Share</TH></tr></thead>
      <tbody>{rows.map((r, i) => (
        <tr key={i} className="border-b border-[#EEF2F5] last:border-0 hover:bg-[#F1F5F9]/50">
          <TD bold>{r.p.name}</TD><TD>{r.pt.name || 'Unnamed'}</TD>
          <TD align="right" mono>{fmtPct(r.pt.sharePct)}</TD>
          <TD align="right" mono>{fmt(r.pt.payout, { currency: r.p.currency })}</TD>
          <TD align="right" mono bold tone={r.pt.profitShare >= 0 ? 'emerald' : 'rose'}>{fmt(r.pt.profitShare, { currency: r.p.currency })}</TD>
        </tr>
      ))}</tbody>
    </table>
  );
};

const CashflowTable = ({ data }) => (
  <table className="w-full">
    <thead><tr className="border-b border-[#D8E4F8] bg-[#EFF6FF]"><TH>Project</TH><TH align="right" help={GLOSSARY.contractValue}>Contract</TH><TH align="right" help={GLOSSARY.advancePct}>Advance</TH><TH align="right" help={GLOSSARY.bankRetention}>Bank Ret.</TH><TH align="right" help={GLOSSARY.advanceCash}>Adv. Cash</TH><TH align="right" help={GLOSSARY.totalDeductions}>Deductions</TH><TH align="right" help={GLOSSARY.realCash}>Real Cash</TH></tr></thead>
    <tbody>{data.map(({ p, c }) => (
      <tr key={p.id} className="border-b border-[#EEF2F5] last:border-0 hover:bg-[#F1F5F9]/50">
        <TD bold>{p.name || 'Untitled'}</TD>
        <TD align="right" mono>{fmtShort(c.contract, p.currency)}</TD>
        <TD align="right" mono>{fmtShort(c.advanceAmount, p.currency)}</TD>
        <TD align="right" mono tone="rose">{fmtShort(c.bankRetained, p.currency)}</TD>
        <TD align="right" mono tone="amber">{fmtShort(c.advanceCashAvailable, p.currency)}</TD>
        <TD align="right" mono tone="rose">{fmtShort(c.totalDeductions, p.currency)}</TD>
        <TD align="right" mono bold tone={c.realCashAfterDeductions >= 0 ? 'emerald' : 'rose'}>{fmtShort(c.realCashAfterDeductions, p.currency)}</TD>
      </tr>
    ))}</tbody>
  </table>
);

// ============================================================
// DOCUMENT CHECKLIST — used in form for managing project documents
// ============================================================
const DocumentChecklist = ({ documents, onChange, projectId }) => {
  const [expandedId, setExpandedId] = useState(null);
  const [showCustom, setShowCustom] = useState(false);
  const [customName, setCustomName] = useState('');
  const [customCategory, setCustomCategory] = useState('Custom');
  const [uploadError, setUploadError] = useState(null);
  const [uploadingId, setUploadingId] = useState(null);

  const docs = documents || [];

  const updateDoc = (templateId, customId, updates) => {
    const matchKey = templateId ? 'templateId' : 'id';
    const matchVal = templateId || customId;
    const exists = docs.find(d => d[matchKey] === matchVal);
    let next;
    if (exists) {
      next = docs.map(d => (d[matchKey] === matchVal ? { ...d, ...updates } : d));
    } else {
      next = [...docs, { templateId, id: customId, status: 'pending', ...updates }];
    }
    onChange(next);
  };

  const removeDoc = (matchKey, matchVal) => {
    onChange(docs.filter(d => d[matchKey] !== matchVal));
  };

  const addCustomDoc = () => {
    if (!customName.trim()) return;
    const newDoc = {
      id: `custom_${Date.now()}`,
      templateId: null,
      customLabel: customName.trim(),
      category: customCategory,
      status: 'pending'
    };
    onChange([...docs, newDoc]);
    setCustomName('');
    setShowCustom(false);
  };

  // Group docs: templates by category, plus custom docs
  const byCategory = {};
  DOC_CATEGORIES.forEach(cat => { byCategory[cat] = { templates: [], customs: [] }; });
  DOCUMENT_TEMPLATES.forEach(t => {
    if (byCategory[t.category]) byCategory[t.category].templates.push(t);
  });
  docs.filter(d => !d.templateId).forEach(d => {
    const cat = d.category || 'Custom';
    if (byCategory[cat]) byCategory[cat].customs.push(d);
  });

  const statusConfig = {
    required: { label: 'Required', color: 'bg-rose-100 text-rose-700 border-rose-200', icon: AlertTriangle },
    pending: { label: 'Pending', color: 'bg-slate-100 text-slate-700 border-slate-200', icon: FileClock },
    received: { label: 'Received', color: 'bg-[#DBEAFE] text-blue-700 border-[#CBD5E1]', icon: FileCheck },
    expired: { label: 'Expired', color: 'bg-rose-100 text-rose-700 border-rose-200', icon: FileX },
    na: { label: 'N/A', color: 'bg-[#EFF6FF] text-blue-700 border-[#D8E4F8]', icon: X }
  };

  const renderDocRow = (template, customDoc) => {
    const isCustom = !!customDoc;
    const matchKey = isCustom ? 'id' : 'templateId';
    const matchVal = isCustom ? customDoc.id : template.id;
    const stored = docs.find(d => d[matchKey] === matchVal);
    const label = isCustom ? customDoc.customLabel : template.label;
    const status = stored?.status || 'pending';
    const sCfg = statusConfig[status];
    const isExpanded = expandedId === matchVal;
    const hasExpiry = template?.hasExpiry || stored?.expiryDate;
    const critical = template?.critical;

    // Check if expired or expiring
    let expiryWarning = null;
    if (stored?.expiryDate && status === 'received') {
      const days = Math.ceil((new Date(stored.expiryDate) - new Date()) / (1000 * 60 * 60 * 24));
      if (days < 0) expiryWarning = { tone: 'rose', text: `Expired ${Math.abs(days)}d ago` };
      else if (days <= 30) expiryWarning = { tone: 'amber', text: `Expires in ${days}d` };
    }

    return (
      <div key={matchVal} className="border border-[#D8E4F8] rounded-lg overflow-hidden bg-white">
        <button
          type="button"
          onClick={() => setExpandedId(isExpanded ? null : matchVal)}
          className="w-full flex items-center gap-3 px-3 py-2.5 hover:bg-[#F1F5F9] transition-colors text-left"
        >
          <div className={`flex-shrink-0 w-7 h-7 rounded-md border flex items-center justify-center ${sCfg.color}`}>
            <sCfg.icon size={13} strokeWidth={2} />
          </div>
          <div className="flex-1 min-w-0">
            <div className="text-sm font-bold text-slate-900 flex items-center gap-2">
              <span className="truncate">{label}</span>
              {critical && <span className="text-[9px] font-bold px-1.5 py-0.5 bg-rose-50 text-rose-600 border border-rose-200 rounded uppercase tracking-wider flex-shrink-0">Critical</span>}
            </div>
            <div className="text-[11px] text-blue-700 mt-0.5 flex items-center gap-2 flex-wrap">
              <span className={`font-bold ${status === 'received' ? 'text-blue-600' : status === 'expired' ? 'text-rose-600' : ''}`}>{sCfg.label}</span>
              {stored?.fileName && <span className="text-blue-600 flex items-center gap-0.5"><File size={9} />{formatBytes(stored.fileSize)}</span>}
              {stored?.url && <span className="text-blue-600 flex items-center gap-0.5"><ExternalLink size={9} />Linked</span>}
              {expiryWarning && <span className={`font-bold ${expiryWarning.tone === 'rose' ? 'text-rose-600' : 'text-slate-600'}`}>· {expiryWarning.text}</span>}
            </div>
          </div>
          <ChevronDown size={16} className={`text-blue-700 transition-transform flex-shrink-0 ${isExpanded ? 'rotate-180' : ''}`} />
        </button>

        {isExpanded && (
          <div className="border-t border-[#D8E4F8] p-3 bg-[#EFF6FF] space-y-3">
            {/* Status selector */}
            <div>
              <label className="block text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1.5">Status</label>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-1">
                {Object.keys(statusConfig).map(s => (
                  <button
                    key={s}
                    type="button"
                    onClick={() => updateDoc(template?.id || null, customDoc?.id, { status: s, ...(isCustom ? { customLabel: label, category: customDoc.category } : {}) })}
                    className={`px-2 py-1.5 text-[11px] font-bold rounded border transition-all ${status === s ? statusConfig[s].color + ' shadow-sm' : 'bg-white border-[#D8E4F8] text-blue-700 hover:border-slate-400'}`}
                  >
                    {statusConfig[s].label}
                  </button>
                ))}
              </div>
            </div>

            {/* URL link */}
            <div>
              <label className="block text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1.5">Document Link (Google Drive, Dropbox, etc.)</label>
              <input
                type="url"
                value={stored?.url || ''}
                onChange={e => updateDoc(template?.id || null, customDoc?.id, { url: e.target.value, ...(isCustom ? { customLabel: label, category: customDoc.category } : {}) })}
                placeholder="https://drive.google.com/..."
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-900 focus:outline-none focus:border-[#2563EB]"
              />
            </div>

            {/* File upload */}
            <div>
              <label className="block text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1.5">Upload File (max {formatBytes(MAX_FILE_BYTES)})</label>
              {stored?.fileName ? (
                <div className="bg-[#EFF6FF] border border-[#CBD5E1] rounded p-2.5 flex items-center justify-between gap-2 flex-wrap">
                  <div className="flex items-center gap-2 min-w-0 flex-1">
                    <File size={16} className="text-blue-600 flex-shrink-0" />
                    <div className="min-w-0">
                      <div className="text-xs font-bold text-blue-900 truncate">{stored.fileName}</div>
                      <div className="text-[10px] text-blue-700 tabular-nums">{formatBytes(stored.fileSize)} · {stored.fileType?.split('/')[1]?.toUpperCase() || 'FILE'}</div>
                    </div>
                  </div>
                  <div className="flex gap-1 flex-shrink-0">
                    <button
                      type="button"
                      onClick={() => previewDocFile(projectId, template?.id || customDoc?.id)}
                      title="Preview"
                      className="p-1.5 text-blue-700 hover:bg-[#D8E4F8] rounded"
                    ><Eye size={13} /></button>
                    <button
                      type="button"
                      onClick={() => downloadDocFile(projectId, template?.id || customDoc?.id, stored.fileName)}
                      title="Download"
                      className="p-1.5 text-blue-700 hover:bg-[#D8E4F8] rounded"
                    ><Download size={13} /></button>
                    <button
                      type="button"
                      onClick={async () => {
                        await deleteDocFile(projectId, template?.id || customDoc?.id);
                        updateDoc(template?.id || null, customDoc?.id, { fileName: null, fileSize: null, fileType: null, uploadedAt: null, ...(isCustom ? { customLabel: label, category: customDoc.category } : {}) });
                      }}
                      title="Remove file"
                      className="p-1.5 text-rose-600 hover:bg-rose-100 rounded"
                    ><Trash2 size={13} /></button>
                  </div>
                </div>
              ) : (
                <label className={`flex items-center justify-center gap-2 px-3 py-2.5 text-xs font-bold border-2 border-dashed rounded cursor-pointer transition-colors ${uploadingId === matchVal ? 'bg-slate-50 border-slate-300 text-slate-700' : 'bg-white border-slate-300 text-slate-900 hover:border-[#2563EB]/50 hover:bg-[#F1F5F9]'}`}>
                  {uploadingId === matchVal ? (
                    <>
                      <div className="w-3.5 h-3.5 border-2 border-slate-500 border-t-transparent rounded-full animate-spin" />
                      Uploading...
                    </>
                  ) : (
                    <>
                      <Upload size={14} />
                      Click to upload PDF, image, or doc
                    </>
                  )}
                  <input
                    type="file"
                    className="hidden"
                    accept=".pdf,.jpg,.jpeg,.png,.gif,.webp,.doc,.docx,.xls,.xlsx,.txt"
                    onChange={async (e) => {
                      const file = e.target.files?.[0];
                      if (!file) return;
                      setUploadError(null);
                      setUploadingId(matchVal);
                      try {
                        const meta = await uploadDocFile(projectId, template?.id || customDoc?.id, file);
                        updateDoc(template?.id || null, customDoc?.id, { ...meta, status: stored?.status === 'received' ? 'received' : 'received', ...(isCustom ? { customLabel: label, category: customDoc.category } : {}) });
                      } catch (err) {
                        setUploadError({ id: matchVal, msg: err.message });
                      } finally {
                        setUploadingId(null);
                        e.target.value = '';
                      }
                    }}
                  />
                </label>
              )}
              {uploadError?.id === matchVal && (
                <div className="mt-2 bg-rose-50 border border-rose-200 text-rose-700 text-[11px] rounded p-2 flex items-start gap-1.5">
                  <AlertTriangle size={12} className="flex-shrink-0 mt-0.5" />
                  <span>{uploadError.msg}</span>
                </div>
              )}
            </div>

            {/* Dates */}
            {(hasExpiry || template?.hasExpiry) && (
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1.5">Issue Date</label>
                  <input
                    type="date"
                    value={stored?.issuedDate || ''}
                    onChange={e => updateDoc(template?.id || null, customDoc?.id, { issuedDate: e.target.value, ...(isCustom ? { customLabel: label, category: customDoc.category } : {}) })}
                    className="w-full bg-white border border-slate-300 rounded px-2 py-1.5 text-xs text-slate-900 focus:outline-none focus:border-[#2563EB]"
                  />
                </div>
                <div>
                  <label className="block text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1.5">Expiry Date</label>
                  <input
                    type="date"
                    value={stored?.expiryDate || ''}
                    onChange={e => updateDoc(template?.id || null, customDoc?.id, { expiryDate: e.target.value, ...(isCustom ? { customLabel: label, category: customDoc.category } : {}) })}
                    className="w-full bg-white border border-slate-300 rounded px-2 py-1.5 text-xs text-slate-900 focus:outline-none focus:border-[#2563EB]"
                  />
                </div>
              </div>
            )}

            {/* Notes */}
            <div>
              <label className="block text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold mb-1.5">Notes</label>
              <textarea
                value={stored?.notes || ''}
                onChange={e => updateDoc(template?.id || null, customDoc?.id, { notes: e.target.value, ...(isCustom ? { customLabel: label, category: customDoc.category } : {}) })}
                rows={2}
                placeholder="Issuer, reference number, conditions..."
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-900 focus:outline-none focus:border-[#2563EB] resize-none"
              />
            </div>

            {/* Actions */}
            <div className="flex gap-2 justify-end">
              {stored?.url && (
                <a href={stored.url} target="_blank" rel="noopener noreferrer" className="inline-flex items-center gap-1.5 text-[11px] font-bold px-2.5 py-1.5 border border-[#2563EB]/30 text-blue-700 bg-white rounded hover:bg-[#F1F5F9]">
                  <ExternalLink size={12} /> Open
                </a>
              )}
              {isCustom && (
                <button type="button" onClick={() => removeDoc('id', customDoc.id)} className="inline-flex items-center gap-1.5 text-[11px] font-bold px-2.5 py-1.5 border border-rose-300 text-rose-600 bg-white rounded hover:bg-rose-50">
                  <Trash2 size={12} /> Remove
                </button>
              )}
              {stored && !isCustom && (
                <button type="button" onClick={() => removeDoc('templateId', template.id)} className="inline-flex items-center gap-1.5 text-[11px] font-bold px-2.5 py-1.5 border border-slate-300 text-blue-700 bg-white rounded hover:bg-[#F1F5F9]">
                  Reset
                </button>
              )}
            </div>
          </div>
        )}
      </div>
    );
  };

  const status = computeDocStatus({ documents: docs });

  return (
    <div>
      {/* Compliance summary */}
      <div className="bg-gradient-to-br from-slate-50 to-white border border-[#D8E4F8] rounded-lg p-3 mb-4">
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <div>
            <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold">Critical Document Compliance</div>
            <div className="text-lg font-bold text-slate-900 tabular-nums">
              {status.criticalReceived.length} / {status.criticalTemplates.length}
              <span className="text-xs text-blue-700 ml-2 font-semibold">({status.compliancePct.toFixed(0)}%)</span>
            </div>
          </div>
          <div className="flex gap-2 text-xs flex-wrap">
            {status.expired.length > 0 && (
              <div className="bg-rose-50 border border-rose-200 text-rose-700 px-2 py-1 rounded font-bold flex items-center gap-1">
                <FileX size={12} /> {status.expired.length} expired
              </div>
            )}
            {status.expiringSoon.length > 0 && (
              <div className="bg-slate-50 border border-slate-200 text-slate-700 px-2 py-1 rounded font-bold flex items-center gap-1">
                <FileClock size={12} /> {status.expiringSoon.length} expiring soon
              </div>
            )}
          </div>
        </div>
        <div className="mt-2 h-1.5 bg-slate-200 rounded-full overflow-hidden">
          <div className={`h-full rounded-full transition-all ${status.compliancePct >= 80 ? 'bg-[#EFF6FF]0' : status.compliancePct >= 50 ? 'bg-slate-500' : 'bg-rose-500'}`} style={{ width: `${status.compliancePct}%` }} />
        </div>
      </div>

      {/* Categories */}
      <div className="space-y-4">
        {DOC_CATEGORIES.map(cat => {
          const { templates: catTemplates, customs: catCustoms } = byCategory[cat];
          if (catTemplates.length === 0 && catCustoms.length === 0 && cat !== 'Custom') return null;
          if (cat === 'Custom' && catCustoms.length === 0) return null;
          return (
            <div key={cat}>
              <div className="text-[11px] uppercase tracking-wider text-[#2563EB] font-semibold mb-2 flex items-center gap-2">
                {cat}
                <div className="flex-1 h-px bg-slate-200" />
              </div>
              <div className="space-y-2">
                {catTemplates.map(t => renderDocRow(t, null))}
                {catCustoms.map(c => renderDocRow(null, c))}
              </div>
            </div>
          );
        })}
      </div>

      {/* Add custom document */}
      <div className="mt-4 pt-4 border-t border-[#D8E4F8]">
        {!showCustom ? (
          <button
            type="button"
            onClick={() => setShowCustom(true)}
            className="w-full inline-flex items-center justify-center gap-2 px-3 py-2 text-xs font-bold text-slate-900 bg-white border border-dashed border-slate-300 rounded-lg hover:border-[#2563EB]/50 hover:text-[#2563EB] transition-colors"
          >
            <Plus size={14} /> Add custom document
          </button>
        ) : (
          <div className="bg-[#EFF6FF] border border-[#D8E4F8] rounded-lg p-3 space-y-2">
            <input
              type="text"
              value={customName}
              onChange={e => setCustomName(e.target.value)}
              placeholder="Document name (e.g. Site Survey Report)"
              className="w-full bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:border-[#2563EB]"
              autoFocus
            />
            <select
              value={customCategory}
              onChange={e => setCustomCategory(e.target.value)}
              className="w-full bg-white border border-slate-300 rounded px-3 py-2 text-sm focus:outline-none focus:border-[#2563EB]"
            >
              {DOC_CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            <div className="flex gap-2 justify-end">
              <button type="button" onClick={() => { setShowCustom(false); setCustomName(''); }} className="px-3 py-1.5 text-xs font-bold text-blue-700 bg-white border border-slate-300 rounded hover:bg-[#F1F5F9]">Cancel</button>
              <button type="button" onClick={addCustomDoc} disabled={!customName.trim()} className="px-3 py-1.5 text-xs font-bold text-slate-900 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 rounded">Add</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};


const ScenarioTester = ({ project, onClose, onApply }) => {
  const [scenario, setScenario] = useState({
    profitPct: project.profitPct,
    advancePct: project.advancePct,
    bankRetentionPct: project.bankRetentionPct,
    extraExpenses: 0
  });

  const baseCalc = calcProject(project);
  const scenarioProject = {
    ...project,
    profitPct: scenario.profitPct,
    advancePct: scenario.advancePct,
    bankRetentionPct: scenario.bankRetentionPct,
    customDeductions: [
      ...(project.customDeductions || []),
      ...(scenario.extraExpenses > 0 ? [{ id: 'scenario_extra', label: 'Scenario Extra Expenses', amount: scenario.extraExpenses }] : [])
    ]
  };
  const scenarioCalc = calcProject(scenarioProject);

  const reset = () => setScenario({
    profitPct: project.profitPct,
    advancePct: project.advancePct,
    bankRetentionPct: project.bankRetentionPct,
    extraExpenses: 0
  });

  const apply = () => {
    onApply({
      ...project,
      profitPct: scenario.profitPct,
      advancePct: scenario.advancePct,
      bankRetentionPct: scenario.bankRetentionPct,
      customDeductions: scenario.extraExpenses > 0
        ? [...(project.customDeductions || []), { id: `extra_${Date.now()}`, label: 'Added via scenario', amount: scenario.extraExpenses }]
        : project.customDeductions
    });
    onClose();
  };

  const Slider = ({ label, value, min, max, step, suffix, onChange, baseline }) => {
    const changed = Math.abs(value - baseline) > 0.001;
    return (
      <div className="bg-[#EFF6FF] rounded-lg p-4">
        <div className="flex items-center justify-between mb-2">
          <label className="text-sm font-bold text-slate-900">{label}</label>
          <div className="flex items-center gap-2">
            {changed && <span className="text-[10px] text-slate-600 font-bold uppercase tracking-wider">Changed</span>}
            <input
              type="number"
              value={value}
              step={step}
              onChange={e => onChange(e.target.value === '' ? 0 : +e.target.value)}
              className="w-24 text-right bg-white border border-slate-300 rounded px-2 py-1 text-sm font-mono tabular-nums focus:outline-none focus:border-[#2563EB]"
            />
            <span className="text-xs text-blue-700 w-3">{suffix}</span>
          </div>
        </div>
        <input
          type="range"
          min={min}
          max={max}
          step={step}
          value={value}
          onChange={e => onChange(+e.target.value)}
          className="w-full h-1.5 bg-slate-200 rounded-full appearance-none cursor-pointer accent-blue-600"
        />
        <div className="flex justify-between text-[10px] text-blue-700 mt-1 tabular-nums">
          <span>{min}{suffix}</span>
          <span className="text-[#2563EB]">Baseline: {baseline}{suffix}</span>
          <span>{max}{suffix}</span>
        </div>
      </div>
    );
  };

  const delta = (a, b) => {
    const d = a - b;
    if (Math.abs(d) < 1) return null;
    return d;
  };

  return (
    <div className="fixed inset-0 bg-white backdrop-blur-sm z-50 flex items-center justify-center p-4 overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl max-w-4xl w-full my-8 max-h-[95vh] overflow-y-auto">
        <div className="sticky top-0 bg-white border-b border-[#D8E4F8] px-6 py-4 flex items-center justify-between rounded-t-2xl z-10">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-gradient-to-br from-slate-500 to-slate-600 rounded-xl flex items-center justify-center">
              <Sliders size={18} className="text-slate-900" />
            </div>
            <div>
              <h2 className="font-bold text-slate-900 text-lg">Scenario Tester</h2>
              <p className="text-xs text-[#2563EB]">Test what-if changes without saving</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-[#F1F5F9] rounded-lg text-[#2563EB]">
            <X size={18} />
          </button>
        </div>

        <div className="p-6">
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Left: Controls */}
            <div className="space-y-3">
              <div className="flex items-center justify-between mb-2">
                <h3 className="text-sm font-bold text-slate-900">Adjust values</h3>
                <button onClick={reset} className="text-xs text-blue-700 hover:text-[#2563EB] inline-flex items-center gap-1">
                  <RotateCcw size={12} /> Reset
                </button>
              </div>
              <Slider label="Profit margin" value={scenario.profitPct} min={0} max={50} step={0.5} suffix="%" onChange={v => setScenario({ ...scenario, profitPct: v })} baseline={project.profitPct} />
              <Slider label="Advance %" value={scenario.advancePct} min={0} max={50} step={1} suffix="%" onChange={v => setScenario({ ...scenario, advancePct: v })} baseline={project.advancePct} />
              <Slider label="Bank retention" value={scenario.bankRetentionPct} min={0} max={30} step={1} suffix="%" onChange={v => setScenario({ ...scenario, bankRetentionPct: v })} baseline={project.bankRetentionPct} />
              <div className="bg-[#EFF6FF] rounded-lg p-4">
                <div className="flex items-center justify-between mb-2">
                  <label className="text-sm font-bold text-slate-900">Extra expenses</label>
                  <input
                    type="number"
                    value={scenario.extraExpenses}
                    onChange={e => setScenario({ ...scenario, extraExpenses: e.target.value === '' ? 0 : +e.target.value })}
                    className="w-32 text-right bg-white border border-slate-300 rounded px-2 py-1 text-sm font-mono tabular-nums focus:outline-none focus:border-[#2563EB]"
                  />
                </div>
                <p className="text-[10px] text-[#2563EB]">Additional cost on top of existing deductions ({project.currency})</p>
              </div>
            </div>

            {/* Right: Comparison */}
            <div className="space-y-3">
              <h3 className="text-sm font-bold text-slate-900 mb-2">Impact analysis</h3>
              <ScenarioStat label="Final Profit" baseline={baseCalc.projectedFinalProfit} scenario={scenarioCalc.projectedFinalProfit} currency={project.currency} good={x => x >= 0} />
              <ScenarioStat label="Real Cash Upfront" baseline={baseCalc.realCashAfterDeductions} scenario={scenarioCalc.realCashAfterDeductions} currency={project.currency} good={x => x >= 0} />
              <ScenarioStat label="Total Deductions" baseline={baseCalc.totalDeductions} scenario={scenarioCalc.totalDeductions} currency={project.currency} good={x => x < baseCalc.totalDeductions} invertDelta />
              <ScenarioStat label="Profit Margin" baseline={baseCalc.profitMargin} scenario={scenarioCalc.profitMargin} suffix="%" decimals={2} good={x => x >= 0} />

              <div className="bg-[#EFF6FF] rounded-lg p-4 mt-4">
                <div className="flex items-center justify-between text-sm">
                  <span className="font-bold text-slate-900">Risk level</span>
                  <div className="flex items-center gap-2">
                    <Pill tone={baseCalc.riskLevel}>{baseCalc.riskLevel}</Pill>
                    <ChevronRight size={14} className="text-[#2563EB]" />
                    <Pill tone={scenarioCalc.riskLevel}>{scenarioCalc.riskLevel}</Pill>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className="flex gap-2 justify-end mt-8 pt-6 border-t border-[#D8E4F8]">
            <Btn onClick={onClose} variant="secondary">Discard</Btn>
            <Btn onClick={apply} icon={Save}>Apply changes to project</Btn>
          </div>
        </div>
      </div>
    </div>
  );
};

const ScenarioStat = ({ label, baseline, scenario, currency = '', suffix = '', decimals = 0, good, invertDelta }) => {
  const d = scenario - baseline;
  const isUp = d > 0;
  const fmtVal = (v) => suffix === '%' ? `${v.toFixed(decimals)}%` : fmtShort(v, currency);
  const direction = invertDelta ? !isUp : isUp;
  const isPositive = good ? good(scenario) : true;
  return (
    <div className="bg-white border border-[#CBD5E1] rounded-lg p-4">
      <div className="text-xs uppercase tracking-wider text-[#2563EB] font-semibold mb-2">{label}</div>
      <div className="flex items-baseline justify-between">
        <div className="flex items-baseline gap-3">
          <div>
            <div className="text-[10px] text-blue-700 uppercase tracking-wider">Baseline</div>
            <div className="text-sm font-mono tabular-nums text-[#2563EB]">{fmtVal(baseline)}</div>
          </div>
          <ChevronRight size={14} className="text-blue-700 self-center" />
          <div>
            <div className="text-[10px] text-blue-700 uppercase tracking-wider">Scenario</div>
            <div className={`text-base font-bold font-mono tabular-nums ${isPositive ? 'text-slate-900' : 'text-rose-600'}`}>{fmtVal(scenario)}</div>
          </div>
        </div>
        {Math.abs(d) > 0.01 && (
          <div className={`text-xs font-bold tabular-nums ${direction ? 'text-blue-600' : 'text-rose-600'}`}>
            {isUp ? '↑' : '↓'} {fmtVal(Math.abs(d)).replace('-', '')}
          </div>
        )}
      </div>
    </div>
  );
};

// ============================================================
// AUTH SCREEN
// ============================================================
const AuthScreen = ({ mode, onSetup, onUnlock, onResetData }) => {
  const [pw, setPw] = useState('');
  const [confirm, setConfirm] = useState('');
  const [error, setError] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [showResetConfirm, setShowResetConfirm] = useState(false);

  const handleSubmit = async () => {
    setError('');
    if (mode === 'setup') {
      if (pw.length < 4) { setError('Password must be at least 4 characters'); return; }
      if (pw !== confirm) { setError('Passwords do not match'); return; }
      setSubmitting(true);
      await onSetup(pw);
      setSubmitting(false);
    } else {
      setSubmitting(true);
      const ok = await onUnlock(pw);
      setSubmitting(false);
      if (!ok) { setError('Incorrect password'); setPw(''); }
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-900 to-blue-950 flex items-center justify-center p-4 relative overflow-hidden" style={{ fontFamily: "'Inter', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
      `}</style>
      {/* Background decoration */}
      <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[600px] h-[600px] bg-[#EFF6FF]0/10 rounded-full blur-3xl pointer-events-none" />
      <div className="absolute bottom-0 right-0 w-[400px] h-[400px] bg-blue-600/10 rounded-full blur-3xl pointer-events-none" />

      <div className="w-full max-w-md relative z-10">
        {/* Logo */}
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-16 h-16 bg-gradient-to-br from-blue-400 to-blue-600 rounded-2xl shadow-2xl shadow-[#2563EB]/25 mb-4">
            <Calculator size={28} className="text-slate-900" strokeWidth={2} />
          </div>
          <h1 className="text-3xl font-bold text-slate-900 tracking-tight">Rays InvestCalc</h1>
          <div className="text-xs text-blue-500 font-bold tracking-[0.2em] uppercase mt-1.5">Pro</div>
        </div>

        {/* Auth card */}
        <div className="bg-white rounded-2xl shadow-2xl p-7">
          <div className="flex items-center gap-2 mb-1">
            {mode === 'setup' ? <Shield size={18} className="text-blue-500" /> : <Lock size={18} className="text-blue-500" />}
            <h2 className="text-xl font-bold text-slate-900">
              {mode === 'setup' ? 'Set up access' : 'Welcome back'}
            </h2>
          </div>
          <p className="text-sm text-blue-700 mb-6">
            {mode === 'setup'
              ? 'Create a password to protect your project data. You will need this every time you open the app.'
              : 'Enter your password to continue.'}
          </p>

          <div className="space-y-4">
            <div>
              <label className="block text-xs font-bold text-slate-900 mb-1.5">Password</label>
              <input
                type="password"
                value={pw}
                onChange={e => setPw(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && !submitting && handleSubmit()}
                autoFocus
                placeholder="Enter your password"
                className="w-full bg-white border border-[#CBD5E1] rounded-lg px-3 py-2.5 text-sm text-slate-900 focus:outline-none focus:border-[#2563EB] focus:ring-2 focus:ring-blue-200 transition-all"
              />
            </div>

            {mode === 'setup' && (
              <div>
                <label className="block text-xs font-bold text-slate-900 mb-1.5">Confirm password</label>
                <input
                  type="password"
                  value={confirm}
                  onChange={e => setConfirm(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && !submitting && handleSubmit()}
                  placeholder="Re-enter your password"
                  className="w-full bg-white border border-[#CBD5E1] rounded-lg px-3 py-2.5 text-sm text-slate-900 focus:outline-none focus:border-[#2563EB] focus:ring-2 focus:ring-blue-200 transition-all"
                />
              </div>
            )}

            {error && (
              <div className="bg-rose-50 border border-rose-200 text-rose-700 text-sm rounded-lg px-3 py-2.5 flex items-center gap-2">
                <AlertTriangle size={14} />
                {error}
              </div>
            )}

            <button
              onClick={handleSubmit}
              disabled={submitting}
              className="w-full bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed text-slate-900 font-bold rounded-lg py-3 text-sm transition-all shadow-sm shadow-blue-500/20"
            >
              {submitting ? 'Please wait...' : mode === 'setup' ? 'Create password & enter' : 'Unlock'}
            </button>
          </div>

          {mode === 'login' && (
            <div className="mt-6 pt-6 border-t border-[#EEF2F5] text-center">
              {!showResetConfirm ? (
                <button onClick={() => setShowResetConfirm(true)} className="text-xs text-blue-700 hover:text-rose-600 transition-colors font-bold">
                  Forgot password?
                </button>
              ) : (
                <div>
                  <p className="text-xs text-rose-600 mb-3 font-bold">This will permanently delete all your projects and reset the password. Continue?</p>
                  <div className="flex gap-2 justify-center">
                    <button onClick={() => setShowResetConfirm(false)} className="text-xs text-slate-900 px-3 py-1.5 hover:bg-[#F1F5F9] rounded-md font-bold border border-[#D8E4F8]">
                      Cancel
                    </button>
                    <button onClick={onResetData} className="text-xs text-slate-900 bg-rose-600 hover:bg-rose-700 px-3 py-1.5 rounded-md font-bold">
                      Reset everything
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Footer note */}
        <div className="text-center mt-6 text-xs text-[#697386] leading-relaxed">
          Project Decision Engine
          <div className="text-blue-700 mt-1">Your data syncs across all devices when signed into the same Claude account.</div>
        </div>
      </div>
    </div>
  );
};

// ============================================================
// MAIN APP
// ============================================================
export default function App() {
  const [projects, setProjects] = useState([]);
  const [view, setView] = useState('dashboard');
  const [activeId, setActiveId] = useState(null);
  const [drilldownMetric, setDrilldownMetric] = useState(null);
  const [scenarioId, setScenarioId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [mobileNav, setMobileNav] = useState(false);
  const [authMode, setAuthMode] = useState('checking');
  const [storedHash, setStoredHash] = useState(null);
  const [toast, setToast] = useState(null);
  const [commandOpen, setCommandOpen] = useState(false);
  const [commandQuery, setCommandQuery] = useState('');

  // Cmd+K / Ctrl+K to open command palette
  useEffect(() => {
    const onKey = (e) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setCommandOpen(o => !o);
      }
      if (e.key === 'Escape') {
        setCommandOpen(false);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  useEffect(() => {
    Promise.all([loadProjects(), loadAuthHash()]).then(([ps, hash]) => {
      setProjects(ps);
      setStoredHash(hash);
      setAuthMode(hash ? 'login' : 'setup');
      setLoading(false);
    });
  }, []);

  useEffect(() => {
    if (toast) {
      const t = setTimeout(() => setToast(null), 3000);
      return () => clearTimeout(t);
    }
  }, [toast]);

  const persist = async (ps) => { setProjects(ps); await saveProjects(ps); };

  const handleSetup = async (pw) => {
    const hash = await hashPassword(pw);
    await saveAuthHash(hash);
    setStoredHash(hash);
    setAuthMode('unlocked');
  };

  const handleUnlock = async (pw) => {
    const hash = await hashPassword(pw);
    if (hash === storedHash) {
      setAuthMode('unlocked');
      return true;
    }
    return false;
  };

  const handleLogout = () => {
    setAuthMode('login');
    setView('dashboard');
    setActiveId(null);
    setDrilldownMetric(null);
    setMobileNav(false);
  };

  const handleResetData = async () => {
    await clearAuthHash();
    try { await window.storage.delete(STORAGE_KEY); } catch {}
    setStoredHash(null);
    setProjects([]);
    setAuthMode('setup');
  };

  const handleKpiClick = (metric, payload) => {
    if (metric === 'best') {
      setActiveId(payload);
      setView('detail');
    } else if (metric === 'totalProjects') {
      setView('projects');
    } else {
      setDrilldownMetric(metric);
      setView('drilldown');
    }
  };

  const handleScenarioApply = async (updatedProject) => {
    const next = projects.map(p => p.id === updatedProject.id ? updatedProject : p);
    await persist(next);
    setToast({ type: 'success', message: 'Scenario applied to project' });
  };

  const handleExport = () => {
    const payload = {
      app: 'rays-investcalc',
      version: 1,
      exportedAt: new Date().toISOString(),
      projects: projects
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `rays-investcalc-backup-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    setToast({ type: 'success', message: `Backup exported: ${projects.length} project${projects.length === 1 ? '' : 's'}` });
  };

  const handleImport = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = JSON.parse(evt.target.result);
        const importedProjects = Array.isArray(data) ? data : data.projects;
        if (!Array.isArray(importedProjects)) throw new Error('Invalid format');
        if (!confirm(`Import ${importedProjects.length} project${importedProjects.length === 1 ? '' : 's'}? This will REPLACE all your current projects.`)) {
          e.target.value = '';
          return;
        }
        await persist(importedProjects);
        setToast({ type: 'success', message: `Imported ${importedProjects.length} project${importedProjects.length === 1 ? '' : 's'}` });
      } catch (err) {
        setToast({ type: 'error', message: 'Import failed: invalid backup file' });
      }
      e.target.value = '';
    };
    reader.readAsText(file);
  };

  const handleSave = async (project) => {
    const exists = projects.some(p => p.id === project.id);
    const next = exists ? projects.map(p => p.id === project.id ? project : p) : [...projects, project];
    await persist(next);
    setView('detail');
    setActiveId(project.id);
  };

  const handleDelete = async (id) => {
    // Clean up any uploaded document files for this project
    const project = projects.find(p => p.id === id);
    if (project?.documents) {
      for (const doc of project.documents) {
        if (doc.fileName) {
          await deleteDocFile(id, doc.templateId || doc.id);
        }
      }
    }
    await persist(projects.filter(p => p.id !== id));
    setConfirmDelete(null); setView('projects'); setActiveId(null);
  };

  const handleDuplicate = async (id) => {
    const orig = projects.find(p => p.id === id);
    if (!orig) return;
    const dup = { ...orig, id: `p_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`, name: `${orig.name} (copy)`, createdAt: new Date().toISOString() };
    await persist([...projects, dup]);
    setActiveId(dup.id); setView('detail');
  };

  const activeProject = projects.find(p => p.id === activeId);

  if (loading) {
    return (
      <div className="min-h-screen bg-[#EFF6FF] flex items-center justify-center">
        <div className="flex items-center gap-3 text-[#2563EB]">
          <div className="w-5 h-5 border-2 border-blue-500 border-t-transparent rounded-full animate-spin" />
          <span className="text-sm font-bold">Loading...</span>
        </div>
      </div>
    );
  }

  if (authMode !== 'unlocked') {
    return <AuthScreen mode={authMode} onSetup={handleSetup} onUnlock={handleUnlock} onResetData={handleResetData} />;
  }

  return (
    <div className="min-h-screen relative" style={{ fontFamily: "'Inter', system-ui, sans-serif", color: '#0F172A' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=Manrope:wght@500;600;700;800;900&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
        body {
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, system-ui, sans-serif;
          font-feature-settings: 'cv11', 'ss01', 'cv01';
          background-color: #F8FAFC;
          background-image: radial-gradient(circle at 1px 1px, rgba(37, 99, 235, 0.06) 1px, transparent 0);
          background-size: 28px 28px;
          color: #0F172A;
          font-weight: 500;
          min-height: 100vh;
        }
        .font-display { font-family: 'Manrope', system-ui, sans-serif; letter-spacing: -0.02em; font-weight: 800; }
        .font-mono { font-family: 'JetBrains Mono', Consolas, monospace; }
        .tabular-nums { font-variant-numeric: tabular-nums; }
        .num-hero { font-family: 'Manrope', system-ui, sans-serif; font-weight: 800; letter-spacing: -0.03em; font-variant-numeric: tabular-nums; color: #071739; }
        .num-display { font-family: 'Manrope', system-ui, sans-serif; font-weight: 800; font-variant-numeric: tabular-nums; letter-spacing: -0.035em; color: #071739; }
        input[type="date"]::-webkit-calendar-picker-indicator { cursor: pointer; opacity: 0.7; }
        input[type="date"]::-webkit-calendar-picker-indicator:hover { opacity: 1; }
        ::-webkit-scrollbar { height: 10px; width: 10px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 5px; }
        ::-webkit-scrollbar-thumb:hover { background: #64748B; }
        .card-shadow { box-shadow: 0 10px 28px rgba(37, 99, 235, 0.08); }
        .card-shadow-lg { box-shadow: 0 18px 45px rgba(37, 99, 235, 0.10); }
        .ink-shadow { box-shadow: 0 1px 0 0 rgba(15,23,42,0.10), 0 4px 14px -2px rgba(15,23,42,0.15); }
        .blue-shadow { box-shadow: 0 8px 20px rgba(37, 99, 235, 0.25); }
        .lift { transition: transform 0.18s cubic-bezier(0.4, 0, 0.2, 1), box-shadow 0.18s cubic-bezier(0.4, 0, 0.2, 1), border-color 0.18s; }
        .lift:hover { transform: translateY(-2px); box-shadow: 0 16px 36px rgba(37, 99, 235, 0.14); }
        @keyframes subtle-pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.5; } }
        .pulse-soft { animation: subtle-pulse 2s ease-in-out infinite; }
      `}</style>

      <div className="flex">
        <Sidebar view={view} setView={(v) => { setView(v); if (v !== 'detail') setActiveId(null); }} projectCount={projects.length} mobileOpen={mobileNav} setMobileOpen={setMobileNav} onLogout={handleLogout} onExport={handleExport} onImport={handleImport} />

        <div className="flex-1 min-w-0">
          {/* Mobile topbar */}
          <div className="lg:hidden bg-white/95 backdrop-blur-xl border-b border-[#D8E4F8] px-4 py-3 flex items-center justify-between sticky top-0 z-30">
            <button onClick={() => setMobileNav(true)} className="p-2 hover:bg-[#F1F5F9] rounded-lg text-slate-700">
              <svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M3 5h14M3 10h14M3 15h14" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/></svg>
            </button>
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 bg-gradient-to-br from-blue-400 to-blue-600 rounded-lg flex items-center justify-center">
                <Calculator size={15} className="text-slate-900" />
              </div>
              <span className="font-bold text-sm text-slate-900">Rays InvestCalc</span>
            </div>
            <button onClick={() => setCommandOpen(true)} className="p-2 hover:bg-[#F1F5F9] rounded-lg text-slate-700">
              <Search size={18} />
            </button>
          </div>

          {/* Desktop topbar with search */}
          <div className="hidden lg:flex items-center justify-between px-8 pt-6 pb-2 sticky top-0 z-30 backdrop-blur-xl">
            <div className="flex-1" />
            <button
              onClick={() => setCommandOpen(true)}
              className="flex items-center gap-3 bg-white hover:bg-[#F1F5F9] border border-[#D8E4F8] hover:border-[#CBD5E1] rounded-lg px-3 py-1.5 text-sm transition-all w-72 text-left"
            >
              <Search size={14} className="text-[#2563EB]" />
              <span className="text-blue-700 flex-1">Search projects, actions...</span>
              <kbd className="text-[10px] font-mono text-blue-700 bg-[#EFF6FF] border border-[#D8E4F8] rounded px-1.5 py-0.5">⌘K</kbd>
            </button>
          </div>

          <main className="p-4 sm:p-6 lg:p-8 lg:pt-2 pb-24 lg:pb-8 max-w-7xl mx-auto">
            {view === 'dashboard' && <Dashboard projects={projects} onNew={() => { setActiveId(null); setView('form'); }} onKpiClick={handleKpiClick} onView={id => { setActiveId(id); setView('detail'); }} />}
            {view === 'projects' && <ProjectsList projects={projects} onView={id => { setActiveId(id); setView('detail'); }} onNew={() => { setActiveId(null); setView('form'); }} />}
            {view === 'drilldown' && drilldownMetric && <Drilldown metric={drilldownMetric} projects={projects} onBack={() => { setView('dashboard'); setDrilldownMetric(null); }} onView={id => { setActiveId(id); setView('detail'); }} />}
            {view === 'form' && <ProjectForm project={activeProject} onSave={handleSave} onCancel={() => setView(activeProject ? 'detail' : 'projects')} />}
            {view === 'detail' && activeProject && <ProjectDetail project={activeProject} onBack={() => setView('projects')} onEdit={(id) => { setActiveId(id); setView('form'); }} onDelete={(id) => setConfirmDelete(id)} onDuplicate={handleDuplicate} onScenario={(id) => setScenarioId(id)} />}
            {view === 'compare' && <Compare projects={projects} />}
            {view === 'reports' && <Reports projects={projects} />}
          </main>

          {/* Mobile floating bottom navigation - premium spec */}
          <nav className="lg:hidden fixed bottom-0 left-0 right-0 z-40 bg-white" style={{
            borderTop: '1px solid #E2E8F0',
            boxShadow: '0 -8px 30px rgba(15, 23, 42, 0.08)'
          }}>
            <div className="grid grid-cols-5 gap-0.5 px-1 py-1.5 safe-area-inset-bottom">
              {[
                { id: 'dashboard', icon: LayoutDashboard, label: 'Home' },
                { id: 'projects', icon: FolderKanban, label: 'Projects' },
                { id: 'form', icon: Plus, label: 'New', isAction: true },
                { id: 'compare', icon: GitCompare, label: 'Compare' },
                { id: 'reports', icon: FileText, label: 'Reports' }
              ].map(item => {
                const isActive = view === item.id || (item.id === 'projects' && view === 'detail');
                if (item.isAction) {
                  return (
                    <button
                      key={item.id}
                      onClick={() => { setActiveId(null); setView('form'); }}
                      className="flex flex-col items-center justify-center py-2 mx-1 -mt-3 text-white"
                      style={{
                        background: 'linear-gradient(135deg, #3B82F6 0%, #2563EB 50%, #1D4ED8 100%)',
                        borderRadius: '18px',
                        boxShadow: '0 12px 30px rgba(37, 99, 235, 0.35)'
                      }}
                    >
                      <item.icon size={22} strokeWidth={2.6} />
                      <span className="text-[9px] font-extrabold mt-0.5">{item.label}</span>
                    </button>
                  );
                }
                return (
                  <button
                    key={item.id}
                    onClick={() => { setView(item.id); setActiveId(null); }}
                    className="flex flex-col items-center justify-center py-2 rounded-lg transition-colors"
                    style={{ color: isActive ? '#2563EB' : '#64748B' }}
                  >
                    <item.icon size={18} strokeWidth={isActive ? 2.6 : 2} />
                    <span className={`text-[9px] mt-0.5 ${isActive ? 'font-extrabold' : 'font-bold'}`}>{item.label}</span>
                  </button>
                );
              })}
            </div>
          </nav>
        </div>
      </div>

      {scenarioId && (() => {
        const sp = projects.find(p => p.id === scenarioId);
        if (!sp) return null;
        return <ScenarioTester project={sp} onClose={() => setScenarioId(null)} onApply={handleScenarioApply} />;
      })()}

      {/* Command Palette (Cmd+K / Ctrl+K) */}
      {commandOpen && (() => {
        const q = commandQuery.toLowerCase().trim();
        const projMatches = projects.filter(p =>
          !q || (p.name || '').toLowerCase().includes(q) || (p.client || '').toLowerCase().includes(q) || (p.location || '').toLowerCase().includes(q)
        ).slice(0, 8);
        const actions = [
          { id: 'nav-dashboard', label: 'Go to Dashboard', icon: LayoutDashboard, do: () => setView('dashboard') },
          { id: 'nav-projects', label: 'Go to Projects', icon: FolderKanban, do: () => setView('projects') },
          { id: 'nav-compare', label: 'Compare Projects', icon: GitCompare, do: () => setView('compare') },
          { id: 'nav-reports', label: 'Go to Reports', icon: FileText, do: () => setView('reports') },
          { id: 'new', label: 'Create New Project', icon: Plus, do: () => { setActiveId(null); setView('form'); } }
        ].filter(a => !q || a.label.toLowerCase().includes(q));

        return (
          <div className="fixed inset-0 z-50 flex items-start justify-center pt-[15vh] px-4 bg-[#1A1F36]/30 backdrop-blur-sm" onClick={() => { setCommandOpen(false); setCommandQuery(''); }}>
            <div className="w-full max-w-xl bg-white border border-[#CBD5E1] rounded-xl shadow-2xl overflow-hidden" onClick={e => e.stopPropagation()}>
              <div className="flex items-center gap-3 p-4 border-b border-[#D8E4F8]">
                <Search size={18} className="text-blue-700 flex-shrink-0" />
                <input
                  autoFocus
                  value={commandQuery}
                  onChange={e => setCommandQuery(e.target.value)}
                  placeholder="Search projects or actions..."
                  className="flex-1 bg-transparent text-slate-900 text-sm placeholder:text-blue-700 focus:outline-none"
                />
                <kbd className="hidden sm:block text-[10px] font-mono text-blue-700 bg-[#EFF6FF] border border-[#D8E4F8] rounded px-1.5 py-0.5">ESC</kbd>
              </div>

              <div className="max-h-[60vh] overflow-y-auto p-2">
                {projMatches.length > 0 && (
                  <div className="mb-2">
                    <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold px-3 py-1.5">Projects</div>
                    {projMatches.map(p => {
                      const c = calcProject(p);
                      return (
                        <button
                          key={p.id}
                          onClick={() => { setActiveId(p.id); setView('detail'); setCommandOpen(false); setCommandQuery(''); }}
                          className="w-full text-left flex items-center gap-3 px-3 py-2.5 rounded-lg hover:bg-[#F1F5F9] group"
                        >
                          <div className="w-8 h-8 rounded-lg bg-[#EFF6FF]0/15 ring-1 ring-blue-200 flex items-center justify-center flex-shrink-0">
                            <Building2 size={14} className="text-blue-500" />
                          </div>
                          <div className="flex-1 min-w-0">
                            <div className="text-sm font-bold text-slate-900 truncate">{p.name || 'Untitled'}</div>
                            <div className="text-xs text-blue-700 truncate">{p.client || 'No client'} · {fmtShort(c.contract, p.currency)}</div>
                          </div>
                          <ChevronRight size={14} className="text-slate-700 group-hover:text-blue-500 transition-colors" />
                        </button>
                      );
                    })}
                  </div>
                )}

                {actions.length > 0 && (
                  <div>
                    <div className="text-[10px] uppercase tracking-wider text-[#2563EB] font-semibold px-3 py-1.5">Actions</div>
                    {actions.map(a => (
                      <button
                        key={a.id}
                        onClick={() => { a.do(); setCommandOpen(false); setCommandQuery(''); }}
                        className="w-full text-left flex items-center gap-3 px-3 py-2.5 rounded-lg hover:bg-[#F1F5F9] group"
                      >
                        <div className="w-8 h-8 rounded-lg bg-[#EFF6FF] flex items-center justify-center flex-shrink-0">
                          <a.icon size={14} className="text-slate-700" />
                        </div>
                        <div className="text-sm font-bold text-slate-900 flex-1">{a.label}</div>
                        <ChevronRight size={14} className="text-slate-700 group-hover:text-blue-500 transition-colors" />
                      </button>
                    ))}
                  </div>
                )}

                {projMatches.length === 0 && actions.length === 0 && (
                  <div className="text-center py-8 text-sm text-[#2563EB]">No matches for "{commandQuery}"</div>
                )}
              </div>

              <div className="border-t border-[#D8E4F8] px-4 py-2 flex items-center justify-between text-[10px] text-[#2563EB]">
                <span>↑↓ navigate · ↵ select</span>
                <span><kbd className="font-mono bg-[#EFF6FF] border border-[#D8E4F8] rounded px-1">Cmd+K</kbd> to toggle</span>
              </div>
            </div>
          </div>
        );
      })()}

      {toast && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 z-50 animate-in fade-in slide-in-from-bottom-2">
          <div className={`${toast.type === 'error' ? 'bg-rose-600' : 'bg-white'} text-slate-900 px-4 py-3 rounded-lg shadow-2xl flex items-center gap-2 text-sm font-bold`}>
            {toast.type === 'error' ? <XCircle size={16} /> : <CheckCircle2 size={16} className="text-blue-500" />}
            {toast.message}
          </div>
        </div>
      )}

      {confirmDelete && (
        <div className="fixed inset-0 bg-white/50 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-xl max-w-md w-full p-6">
            <div className="flex items-start gap-4 mb-4">
              <div className="w-10 h-10 bg-rose-100 rounded-full flex items-center justify-center flex-shrink-0">
                <AlertTriangle className="text-rose-600" size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-900 text-lg">Delete project?</h3>
                <p className="text-sm text-blue-700 mt-1">This cannot be undone. All calculations and notes will be permanently removed.</p>
              </div>
            </div>
            <div className="flex gap-2 justify-end">
              <Btn onClick={() => setConfirmDelete(null)} variant="secondary">Cancel</Btn>
              <Btn onClick={() => handleDelete(confirmDelete)} variant="dangerSolid" icon={Trash2}>Delete Project</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
