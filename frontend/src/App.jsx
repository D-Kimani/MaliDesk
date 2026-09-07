
import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  Search, Bell, Settings, ChevronRight, ChevronDown, ChevronUp, Plus, X,
  ArrowLeft, Printer, AlertTriangle, Wallet, Users, DoorClosed, CircleDollarSign,
  Home, ListChecks, Archive as ArchiveIcon, UploadCloud, Download, FileSpreadsheet,
  FileDown, Circle, Filter, HelpCircle, MoreVertical, Pencil, Trash2, Lock, LogOut, ShieldCheck, UserPlus,
} from "lucide-react";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------------------- */
/* Design tokens — MaliDesk palette                                        */
/* ---------------------------------------------------------------------- */
const BEIGE = "#F8F5F0";
const SIDEBAR = "#1A3C34";
const SIDEBAR_ACTIVE = "#234B41";
const SIDEBAR_BORDER = "#2B5347";
const AMBER = "#D9A441";
const AMBER_DARK = "#B7841F";
const CARD = "#FFFFFF";
const INK = "#26312D";
const MUTED = "#7A8880";
const BORDER = "#EAE3D6";

const STATUS_COLORS = {
  VACANT: { bg: "#FFF3D9", fg: "#92650B", dot: "#E8B846" },
  PAID: { bg: "#E3F5EA", fg: "#1F8A4C", dot: "#3FB873" },
  LOW: { bg: "#FDF1D9", fg: "#92650B", dot: "#D9A441" },
  MEDIUM: { bg: "#FCE6D2", fg: "#B4530A", dot: "#EF9346" },
  HIGH: { bg: "#FBE3E3", fg: "#C0392B", dot: "#E06666" },
};

const kes = (n) =>
  new Intl.NumberFormat("en-KE", { style: "currency", currency: "KES", maximumFractionDigits: 0 }).format(
    Number.isFinite(n) ? n : 0
  );
const fmtDate = (d) => (d ? new Date(d).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "—");
const monthLabel = (ym) => {
  if (!ym) return "—";
  const [y, m] = ym.split("-");
  return new Date(Number(y), Number(m) - 1, 1).toLocaleDateString("en-GB", { month: "short", year: "numeric" });
};
const uid = (p) => {
  try {
    if (globalThis.crypto?.randomUUID) return `${p}_${globalThis.crypto.randomUUID().replace(/-/g, "").slice(0, 16)}`;
  } catch {}
  return `${p}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
};
const DATA_KEY = "malidesk-data-v3";
const LEGACY_DATA_KEY = "malidesk-data-v1";
const DATA_VERSION = 3;
const AUTH_API_BASE = "https://malidesk.onrender.com";
// SECURITY NOTE: the existing local data layer is intentionally preserved for offline compatibility.
// Production deployments must route sensitive mutations through the authenticated API layer.

const todayISO = () => new Date().toISOString().slice(0, 10);
const monthOf = (dateStr) => (dateStr || todayISO()).slice(0, 7);

const normaliseData = (raw) => {
  const d = raw && typeof raw === "object" ? raw : {};
  // Transactions with no date can't be sorted or placed on the ledger, and
  // shouldn't exist going forward — but a handful may already be sitting in
  // saved data from an earlier import. Drop them here so every load is safe.
  const cleanTxns = (Array.isArray(d.transactions) ? d.transactions : []).filter(
    (t) => t && typeof t.date === "string" && t.date.trim() !== ""
  );
  return {
    version: DATA_VERSION,
    units: Array.isArray(d.units) ? d.units : [],
    tenants: Array.isArray(d.tenants) ? d.tenants : [],
    transactions: cleanTxns,
    archive: Array.isArray(d.archive) ? d.archive : [],
    auditLog: Array.isArray(d.auditLog) ? d.auditLog : [],
    settings: d.settings && typeof d.settings === "object" ? d.settings : {},
    // Content-based fingerprints of transactions the user has deliberately
    // deleted. Keyed by unit CODE (not internal id, which is regenerated on
    // every full-replace import) so that re-importing a worksheet that still
    // contains the old row never silently resurrects it.
    deletedImportKeys: Array.isArray(d.deletedImportKeys) ? d.deletedImportKeys : [],
  };
};

// Stable content fingerprint for a transaction, used to (a) recognise the
// "same" row across repeated imports of the same worksheet, and (b) let a
// deletion "stick" through a later re-import. Built from the unit CODE
// (stable) rather than the internal unitId (regenerated every import), plus
// the fields that identify a real-world charge or payment.
const txImportKey = (unitCode, t) =>
  [
    String(unitCode || "").trim().toUpperCase(),
    t.type,
    String(t.date || ""),
    Math.round((Number(t.rentCharged) || 0) * 100),
    Math.round((Number(t.garbageCharged) || 0) * 100),
    Math.round((Number(t.paymentReceived) || 0) * 100),
    String(t.description || "").trim().toLowerCase(),
    String(t.mpesaRef || "").trim().toUpperCase(),
  ].join("|");

const storageGet = async (key) => {
  try {
    if (window.storage?.get) {
      const res = await window.storage.get(key, false);
      if (res?.value) return JSON.parse(res.value);
    }
  } catch {}
  return null;
};

const storageSet = async (key, value) => {
  const serialised = JSON.stringify(value);
  try {
    if (window.storage?.set) {
      await window.storage.set(key, serialised, false);
      return true;
    }
  } catch {}
  return false;
};

/* ---------------------------------------------------------------------- */
/* Seed data — your real workbook data                                     */
/* ---------------------------------------------------------------------- */
function seedData() {
  return {"units":[{"id":"u_1","code":"A1","type":"residential","rent":5000,"garbage":100,"depositHeld":5000,"currentTenantId":"t_2","occupancyStatus":"occupied"},{"id":"u_15","code":"A2","type":"residential","rent":4500,"garbage":0,"depositHeld":0,"currentTenantId":"t_16","occupancyStatus":"occupied"},{"id":"u_73","code":"A3","type":"residential","rent":4500,"garbage":200,"depositHeld":4500,"currentTenantId":"t_74","occupancyStatus":"occupied"},{"id":"u_91","code":"A4","type":"residential","rent":3000,"garbage":200,"depositHeld":0,"currentTenantId":"t_92","occupancyStatus":"occupied"},{"id":"u_149","code":"A5","type":"residential","rent":3000,"garbage":200,"depositHeld":0,"currentTenantId":"t_150","occupancyStatus":"occupied"},{"id":"u_165","code":"A6","type":"residential","rent":3000,"garbage":200,"depositHeld":0,"currentTenantId":"t_166","occupancyStatus":"occupied"},{"id":"u_190","code":"A7","type":"residential","rent":3000,"garbage":100,"depositHeld":0,"currentTenantId":"t_191","occupancyStatus":"occupied"},{"id":"u_242","code":"B1","type":"residential","rent":3000,"garbage":200,"depositHeld":0,"currentTenantId":"t_243","occupancyStatus":"occupied"},{"id":"u_287","code":"B2","type":"residential","rent":6000,"garbage":200,"depositHeld":1000,"currentTenantId":"t_288","occupancyStatus":"occupied"},{"id":"u_313","code":"B3","type":"residential","rent":4500,"garbage":200,"depositHeld":0,"currentTenantId":"t_314","occupancyStatus":"occupied"},{"id":"u_366","code":"C1","type":"residential","rent":7500,"garbage":200,"depositHeld":0,"currentTenantId":"t_367","occupancyStatus":"occupied"},{"id":"u_420","code":"C2","type":"residential","rent":3000,"garbage":200,"depositHeld":1000,"currentTenantId":"t_421","occupancyStatus":"occupied"},{"id":"u_440","code":"C3","type":"residential","rent":5500,"garbage":200,"depositHeld":0,"currentTenantId":"t_441","occupancyStatus":"occupied"},{"id":"u_517","code":"Shop 1","type":"shop","rent":10000,"garbage":0,"depositHeld":0,"currentTenantId":"t_518","occupancyStatus":"occupied"},{"id":"u_542","code":"Shop 2","type":"shop","rent":10000,"garbage":0,"depositHeld":10000,"currentTenantId":"t_543","occupancyStatus":"occupied"}],"tenants":[{"id":"t_2","name":"Colins Lugua","phone":"","email":"","moveInDate":"2026-03-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_16","name":"KATIWA KAMENE","phone":"","email":"","moveInDate":"2024-05-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_74","name":"Shaline Chepkirui/HILARY","phone":"","email":"","moveInDate":"2026-01-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_92","name":"RAPHAEL NDELEVA / BENEDETTAH NDELEVA","phone":"","email":"","moveInDate":"2024-05-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_150","name":"Joseph Musyoka","phone":"","email":"","moveInDate":"2026-02-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_166","name":"SAMUEL VICTOR BEET","phone":"","email":"","moveInDate":"2025-11-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_191","name":"ODENY ODONGO","phone":"","email":"","moveInDate":"2024-05-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_243","name":"JACKLYNE CHENINGINGWA","phone":"","email":"","moveInDate":"2024-10-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_288","name":"ANITA NYAKUNDI / DAVID ODHIAMBO","phone":"","email":"","moveInDate":"2025-09-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_314","name":"FRANKLIN JOHN","phone":"","email":"","moveInDate":"2024-05-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_367","name":"HANINGTON WACIRA / POLLY","phone":"","email":"","moveInDate":"2024-04-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_421","name":"DEDAN Kingori/SAMUEL KIMATHI/ARON MOROMPI","phone":"","email":"","moveInDate":"2026-03-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_441","name":"MILLICENT ODERA","phone":"","email":"","moveInDate":"2024-05-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_518","name":"Allan Macharia /Ann Mwangi","phone":"","email":"","moveInDate":"2025-08-01","moveOutDate":null,"notes":"Imported from Excel."},{"id":"t_543","name":"Mary Macharia","phone":"","email":"","moveInDate":"2025-10-01","moveOutDate":null,"notes":"Imported from Excel."}],"transactions":[{"id":"tx_3","unitId":"u_1","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"5,000 Depo","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_4","unitId":"u_1","date":"2026-03-18","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4900,"mpesaRef":"UCIF49T5M4","remarks":"","createdBy":"Import","createdAt":"2026-03-18"},{"id":"tx_5","unitId":"u_1","date":"2026-03-22","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2500,"mpesaRef":"UCMF4A800A","remarks":"Depo cleared","createdBy":"Import","createdAt":"2026-03-22"},{"id":"tx_6","unitId":"u_1","date":"2026-03-24","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"UCOF4AG7GA","remarks":"","createdBy":"Import","createdAt":"2026-03-24"},{"id":"tx_7","unitId":"u_1","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":5000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_8","unitId":"u_1","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":5000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_9","unitId":"u_1","date":"2026-05-04","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"UE4F437KW2","remarks":"","createdBy":"Import","createdAt":"2026-05-04"},{"id":"tx_10","unitId":"u_1","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":5000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_11","unitId":"u_1","date":"2026-06-03","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"UF3F46N58H","remarks":"","createdBy":"Import","createdAt":"2026-06-03"},{"id":"tx_12","unitId":"u_1","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":5000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_13","unitId":"u_1","date":"2026-07-03","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5100,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-03"},{"id":"tx_14","unitId":"u_1","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":5000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_17","unitId":"u_15","date":"2024-05-01","month":"2024-05","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-05-01"},{"id":"tx_18","unitId":"u_15","date":"2024-05-05","month":"2024-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"SE57AGDEVX","remarks":"","createdBy":"Import","createdAt":"2024-05-05"},{"id":"tx_19","unitId":"u_15","date":"2024-06-01","month":"2024-06","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-06-01"},{"id":"tx_20","unitId":"u_15","date":"2024-06-05","month":"2024-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"SF54JVEV5Y","remarks":"","createdBy":"Import","createdAt":"2024-06-05"},{"id":"tx_21","unitId":"u_15","date":"2024-07-01","month":"2024-07","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-07-01"},{"id":"tx_22","unitId":"u_15","date":"2024-07-06","month":"2024-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"SG63QGNDV9","remarks":"","createdBy":"Import","createdAt":"2024-07-06"},{"id":"tx_23","unitId":"u_15","date":"2024-08-01","month":"2024-08","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-08-01"},{"id":"tx_24","unitId":"u_15","date":"2024-08-06","month":"2024-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"SH654CXXY9","remarks":"","createdBy":"Import","createdAt":"2024-08-06"},{"id":"tx_25","unitId":"u_15","date":"2024-09-01","month":"2024-09","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-09-01"},{"id":"tx_26","unitId":"u_15","date":"2024-09-07","month":"2024-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"SI72PLCW0I","remarks":"","createdBy":"Import","createdAt":"2024-09-07"},{"id":"tx_27","unitId":"u_15","date":"2024-10-01","month":"2024-10","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-10-01"},{"id":"tx_28","unitId":"u_15","date":"2024-10-06","month":"2024-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"SJ643CA9I2","remarks":"","createdBy":"Import","createdAt":"2024-10-06"},{"id":"tx_29","unitId":"u_15","date":"2024-11-01","month":"2024-11","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-11-01"},{"id":"tx_30","unitId":"u_15","date":"2024-11-06","month":"2024-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"SK67U4HX7V","remarks":"","createdBy":"Import","createdAt":"2024-11-06"},{"id":"tx_31","unitId":"u_15","date":"2024-12-01","month":"2024-12","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-12-01"},{"id":"tx_32","unitId":"u_15","date":"2024-12-05","month":"2024-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"SL59DSH8KN","remarks":"","createdBy":"Import","createdAt":"2024-12-05"},{"id":"tx_33","unitId":"u_15","date":"2025-01-01","month":"2025-01","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-01-01"},{"id":"tx_34","unitId":"u_15","date":"2025-01-06","month":"2025-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TA608ISC04","remarks":"","createdBy":"Import","createdAt":"2025-01-06"},{"id":"tx_35","unitId":"u_15","date":"2025-02-01","month":"2025-02","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-02-01"},{"id":"tx_36","unitId":"u_15","date":"2025-02-06","month":"2025-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6000,"mpesaRef":"TB681U2C3U","remarks":"","createdBy":"Import","createdAt":"2025-02-06"},{"id":"tx_37","unitId":"u_15","date":"2025-03-01","month":"2025-03","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-03-01"},{"id":"tx_38","unitId":"u_15","date":"2025-03-11","month":"2025-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TCB6808HS0","remarks":"","createdBy":"Import","createdAt":"2025-03-11"},{"id":"tx_39","unitId":"u_15","date":"2025-04-01","month":"2025-04","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-04-01"},{"id":"tx_40","unitId":"u_15","date":"2025-04-07","month":"2025-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TD73T0MGB1","remarks":"","createdBy":"Import","createdAt":"2025-04-07"},{"id":"tx_41","unitId":"u_15","date":"2025-05-01","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-01"},{"id":"tx_42","unitId":"u_15","date":"2025-05-05","month":"2025-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TE58KFI578","remarks":"","createdBy":"Import","createdAt":"2025-05-05"},{"id":"tx_43","unitId":"u_15","date":"2025-06-01","month":"2025-06","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-06-01"},{"id":"tx_44","unitId":"u_15","date":"2025-06-07","month":"2025-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TF71YMZMWL","remarks":"","createdBy":"Import","createdAt":"2025-06-07"},{"id":"tx_45","unitId":"u_15","date":"2025-07-01","month":"2025-07","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-07-01"},{"id":"tx_46","unitId":"u_15","date":"2025-07-07","month":"2025-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TG74ZQPORC","remarks":"","createdBy":"Import","createdAt":"2025-07-07"},{"id":"tx_47","unitId":"u_15","date":"2025-08-01","month":"2025-08","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-08-01"},{"id":"tx_48","unitId":"u_15","date":"2025-08-05","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TH503DKKO4","remarks":"","createdBy":"Import","createdAt":"2025-08-05"},{"id":"tx_49","unitId":"u_15","date":"2025-09-01","month":"2025-09","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-09-01"},{"id":"tx_50","unitId":"u_15","date":"2025-09-06","month":"2025-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TI61MELPOJ","remarks":"","createdBy":"Import","createdAt":"2025-09-06"},{"id":"tx_51","unitId":"u_15","date":"2025-10-01","month":"2025-10","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_52","unitId":"u_15","date":"2025-10-05","month":"2025-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TJ55A6NB9I","remarks":"","createdBy":"Import","createdAt":"2025-10-05"},{"id":"tx_53","unitId":"u_15","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_54","unitId":"u_15","date":"2025-11-06","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TK65A9JNEM","remarks":"","createdBy":"Import","createdAt":"2025-11-06"},{"id":"tx_55","unitId":"u_15","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_56","unitId":"u_15","date":"2025-12-09","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"TL95A0N451","remarks":"","createdBy":"Import","createdAt":"2025-12-09"},{"id":"tx_57","unitId":"u_15","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_58","unitId":"u_15","date":"2026-01-07","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UA75A3BV4U","remarks":"","createdBy":"Import","createdAt":"2026-01-07"},{"id":"tx_59","unitId":"u_15","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_60","unitId":"u_15","date":"2026-02-07","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UB75A69STV","remarks":"","createdBy":"Import","createdAt":"2026-02-07"},{"id":"tx_61","unitId":"u_15","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_62","unitId":"u_15","date":"2026-03-07","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6000,"mpesaRef":"UC75A8Z2ZV","remarks":"","createdBy":"Import","createdAt":"2026-03-07"},{"id":"tx_63","unitId":"u_15","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_64","unitId":"u_15","date":"2026-04-06","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UD65A0771S","remarks":"","createdBy":"Import","createdAt":"2026-04-06"},{"id":"tx_65","unitId":"u_15","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_66","unitId":"u_15","date":"2026-05-03","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UE35A3DTAF","remarks":"","createdBy":"Import","createdAt":"2026-05-03"},{"id":"tx_67","unitId":"u_15","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_68","unitId":"u_15","date":"2026-06-07","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UF75A7H5IT","remarks":"","createdBy":"Import","createdAt":"2026-06-07"},{"id":"tx_69","unitId":"u_15","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_70","unitId":"u_15","date":"2026-07-06","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UG65AATKKN4","remarks":"","createdBy":"Import","createdAt":"2026-07-06"},{"id":"tx_71","unitId":"u_15","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_72","unitId":"u_15","date":"2026-08-03","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UH35A27V6D","remarks":"","createdBy":"Import","createdAt":"2026-08-03"},{"id":"tx_75","unitId":"u_73","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"Depo c/f from previous tenant","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_76","unitId":"u_73","date":"2026-01-13","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UADFU3JC98","remarks":"","createdBy":"Import","createdAt":"2026-01-13"},{"id":"tx_77","unitId":"u_73","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_78","unitId":"u_73","date":"2026-02-08","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UB8F264RJM","remarks":"","createdBy":"Import","createdAt":"2026-02-08"},{"id":"tx_79","unitId":"u_73","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_80","unitId":"u_73","date":"2026-03-07","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5100,"mpesaRef":"UC7O98C3B3","remarks":"","createdBy":"Import","createdAt":"2026-03-07"},{"id":"tx_81","unitId":"u_73","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_82","unitId":"u_73","date":"2026-04-18","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4700,"mpesaRef":"UDIFU11SD2","remarks":"","createdBy":"Import","createdAt":"2026-04-18"},{"id":"tx_83","unitId":"u_73","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_84","unitId":"u_73","date":"2026-05-16","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4700,"mpesaRef":"UEGF24H7JL","remarks":"","createdBy":"Import","createdAt":"2026-05-16"},{"id":"tx_85","unitId":"u_73","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_86","unitId":"u_73","date":"2026-06-21","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4700,"mpesaRef":"UF11FU89MKJ","remarks":"","createdBy":"Import","createdAt":"2026-06-21"},{"id":"tx_87","unitId":"u_73","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_88","unitId":"u_73","date":"2026-07-28","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4700,"mpesaRef":"UGSF20TNYT","remarks":"","createdBy":"Import","createdAt":"2026-07-28"},{"id":"tx_89","unitId":"u_73","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_90","unitId":"u_73","date":"2026-08-24","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4700,"mpesaRef":"UHOF23XZXA","remarks":"","createdBy":"Import","createdAt":"2026-08-24"},{"id":"tx_93","unitId":"u_91","date":"2024-05-01","month":"2024-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-05-01"},{"id":"tx_94","unitId":"u_91","date":"2024-05-08","month":"2024-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"SE86JZPCN2","remarks":"","createdBy":"Import","createdAt":"2024-05-08"},{"id":"tx_95","unitId":"u_91","date":"2024-06-01","month":"2024-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-06-01"},{"id":"tx_96","unitId":"u_91","date":"2024-06-09","month":"2024-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"SF96VY6NXY","remarks":"","createdBy":"Import","createdAt":"2024-06-09"},{"id":"tx_97","unitId":"u_91","date":"2024-07-01","month":"2024-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-07-01"},{"id":"tx_98","unitId":"u_91","date":"2024-07-09","month":"2024-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"SG974VW31B","remarks":"","createdBy":"Import","createdAt":"2024-07-09"},{"id":"tx_99","unitId":"u_91","date":"2024-08-01","month":"2024-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-08-01"},{"id":"tx_100","unitId":"u_91","date":"2024-08-10","month":"2024-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"SHA2JVN500","remarks":"","createdBy":"Import","createdAt":"2024-08-10"},{"id":"tx_101","unitId":"u_91","date":"2024-09-01","month":"2024-09","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-09-01"},{"id":"tx_102","unitId":"u_91","date":"2024-09-10","month":"2024-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"SIA92B1A6T","remarks":"","createdBy":"Import","createdAt":"2024-09-10"},{"id":"tx_103","unitId":"u_91","date":"2024-10-01","month":"2024-10","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-10-01"},{"id":"tx_104","unitId":"u_91","date":"2024-10-10","month":"2024-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"SJA5JU1PN7","remarks":"","createdBy":"Import","createdAt":"2024-10-10"},{"id":"tx_105","unitId":"u_91","date":"2024-11-01","month":"2024-11","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-11-01"},{"id":"tx_106","unitId":"u_91","date":"2024-11-10","month":"2024-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"SKA8CNTZV8","remarks":"","createdBy":"Import","createdAt":"2024-11-10"},{"id":"tx_107","unitId":"u_91","date":"2024-12-01","month":"2024-12","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-12-01"},{"id":"tx_108","unitId":"u_91","date":"2024-12-10","month":"2024-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"SLA8ZXHCHE","remarks":"","createdBy":"Import","createdAt":"2024-12-10"},{"id":"tx_109","unitId":"u_91","date":"2025-01-01","month":"2025-01","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-01-01"},{"id":"tx_110","unitId":"u_91","date":"2025-01-09","month":"2025-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TA93LIURJX","remarks":"","createdBy":"Import","createdAt":"2025-01-09"},{"id":"tx_111","unitId":"u_91","date":"2025-02-01","month":"2025-02","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-02-01"},{"id":"tx_112","unitId":"u_91","date":"2025-02-10","month":"2025-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TBA8G3NFOG","remarks":"","createdBy":"Import","createdAt":"2025-02-10"},{"id":"tx_113","unitId":"u_91","date":"2025-03-01","month":"2025-03","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-03-01"},{"id":"tx_114","unitId":"u_91","date":"2025-03-09","month":"2025-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TC931LUOB5","remarks":"","createdBy":"Import","createdAt":"2025-03-09"},{"id":"tx_115","unitId":"u_91","date":"2025-04-01","month":"2025-04","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-04-01"},{"id":"tx_116","unitId":"u_91","date":"2025-04-10","month":"2025-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TDA87E6O5M","remarks":"","createdBy":"Import","createdAt":"2025-04-10"},{"id":"tx_117","unitId":"u_91","date":"2025-05-01","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-01"},{"id":"tx_118","unitId":"u_91","date":"2025-05-10","month":"2025-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TEA06SQNHA","remarks":"","createdBy":"Import","createdAt":"2025-05-10"},{"id":"tx_119","unitId":"u_91","date":"2025-05-02","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-02"},{"id":"tx_120","unitId":"u_91","date":"2025-06-10","month":"2025-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TFA3A1JX4L","remarks":"","createdBy":"Import","createdAt":"2025-06-10"},{"id":"tx_121","unitId":"u_91","date":"2025-05-03","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-03"},{"id":"tx_122","unitId":"u_91","date":"2024-07-10","month":"2024-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TGA1DUSCK3","remarks":"","createdBy":"Import","createdAt":"2024-07-10"},{"id":"tx_123","unitId":"u_91","date":"2025-05-04","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-04"},{"id":"tx_124","unitId":"u_91","date":"2025-08-10","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"THA2SHDDI8","remarks":"","createdBy":"Import","createdAt":"2025-08-10"},{"id":"tx_125","unitId":"u_91","date":"2025-05-05","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-05"},{"id":"tx_126","unitId":"u_91","date":"2025-09-09","month":"2025-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TI9444DD9G","remarks":"","createdBy":"Import","createdAt":"2025-09-09"},{"id":"tx_127","unitId":"u_91","date":"2025-05-06","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-06"},{"id":"tx_128","unitId":"u_91","date":"2025-10-10","month":"2025-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TJA866WD2O","remarks":"","createdBy":"Import","createdAt":"2025-10-10"},{"id":"tx_129","unitId":"u_91","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_130","unitId":"u_91","date":"2025-11-11","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TKB869PPT2","remarks":"","createdBy":"Import","createdAt":"2025-11-11"},{"id":"tx_131","unitId":"u_91","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_132","unitId":"u_91","date":"2025-12-10","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TLA860IG77","remarks":"","createdBy":"Import","createdAt":"2025-12-10"},{"id":"tx_133","unitId":"u_91","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_134","unitId":"u_91","date":"2026-01-11","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UAB863DHXT","remarks":"","createdBy":"Import","createdAt":"2026-01-11"},{"id":"tx_135","unitId":"u_91","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_136","unitId":"u_91","date":"2026-02-10","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UBA8665D1T","remarks":"","createdBy":"Import","createdAt":"2026-02-10"},{"id":"tx_137","unitId":"u_91","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_138","unitId":"u_91","date":"2026-03-10","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UCA868PYRO","remarks":"","createdBy":"Import","createdAt":"2026-03-10"},{"id":"tx_139","unitId":"u_91","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_140","unitId":"u_91","date":"2026-04-11","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UDB8609U2D","remarks":"","createdBy":"Import","createdAt":"2026-04-11"},{"id":"tx_141","unitId":"u_91","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_142","unitId":"u_91","date":"2026-05-09","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UE9863H978","remarks":"","createdBy":"Import","createdAt":"2026-05-09"},{"id":"tx_143","unitId":"u_91","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_144","unitId":"u_91","date":"2026-06-13","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UFD867HYKV","remarks":"","createdBy":"Import","createdAt":"2026-06-13"},{"id":"tx_145","unitId":"u_91","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_146","unitId":"u_91","date":"2026-07-12","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UGCPFAZDD2","remarks":"BENEDETTAH NDELEVA","createdBy":"Import","createdAt":"2026-07-12"},{"id":"tx_147","unitId":"u_91","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_148","unitId":"u_91","date":"2026-08-12","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6400,"mpesaRef":"UHC862HCU","remarks":"","createdBy":"Import","createdAt":"2026-08-12"},{"id":"tx_151","unitId":"u_149","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_152","unitId":"u_149","date":"2026-02-13","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"ubddq6egva","remarks":"","createdBy":"Import","createdAt":"2026-02-13"},{"id":"tx_153","unitId":"u_149","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_154","unitId":"u_149","date":"2026-03-10","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UCADQ8RPFP","remarks":"","createdBy":"Import","createdAt":"2026-03-10"},{"id":"tx_155","unitId":"u_149","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_156","unitId":"u_149","date":"2026-04-11","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UDBDQOANDR","remarks":"","createdBy":"Import","createdAt":"2026-04-11"},{"id":"tx_157","unitId":"u_149","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_158","unitId":"u_149","date":"2026-05-11","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UEBDQ3OJVX","remarks":"","createdBy":"Import","createdAt":"2026-05-11"},{"id":"tx_159","unitId":"u_149","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_160","unitId":"u_149","date":"2026-06-12","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UFCDQ79PMC","remarks":"","createdBy":"Import","createdAt":"2026-06-12"},{"id":"tx_161","unitId":"u_149","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_162","unitId":"u_149","date":"2026-07-13","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UGDDQARLX0","remarks":"","createdBy":"Import","createdAt":"2026-07-13"},{"id":"tx_163","unitId":"u_149","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_164","unitId":"u_149","date":"2026-08-12","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UHCDQ26GRT","remarks":"","createdBy":"Import","createdAt":"2026-08-12"},{"id":"tx_167","unitId":"u_165","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_168","unitId":"u_165","date":"2025-11-21","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TKLA4SAS3T8","remarks":"","createdBy":"Import","createdAt":"2025-11-21"},{"id":"tx_169","unitId":"u_165","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_170","unitId":"u_165","date":"2025-12-03","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"TL34SBU1VB","remarks":"","createdBy":"Import","createdAt":"2025-12-03"},{"id":"tx_171","unitId":"u_165","date":"2025-12-13","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1200,"mpesaRef":"TDL4SOWW8K","remarks":"","createdBy":"Import","createdAt":"2025-12-13"},{"id":"tx_172","unitId":"u_165","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_173","unitId":"u_165","date":"2026-01-06","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1500,"mpesaRef":"UA64S32JO1","remarks":"","createdBy":"Import","createdAt":"2026-01-06"},{"id":"tx_174","unitId":"u_165","date":"2026-01-16","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1700,"mpesaRef":"UAG4S3ZOC3","remarks":"","createdBy":"Import","createdAt":"2026-01-16"},{"id":"tx_175","unitId":"u_165","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_176","unitId":"u_165","date":"2026-02-04","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1700,"mpesaRef":"UB44S5QQGL","remarks":"","createdBy":"Import","createdAt":"2026-02-04"},{"id":"tx_177","unitId":"u_165","date":"2026-02-24","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1500,"mpesaRef":"UBO4S7L8EH","remarks":"","createdBy":"Import","createdAt":"2026-02-24"},{"id":"tx_178","unitId":"u_165","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_179","unitId":"u_165","date":"2026-03-09","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"UC94S8TX9Y","remarks":"","createdBy":"Import","createdAt":"2026-03-09"},{"id":"tx_180","unitId":"u_165","date":"2026-03-23","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1200,"mpesaRef":"UCN4SA7F8S","remarks":"","createdBy":"Import","createdAt":"2026-03-23"},{"id":"tx_181","unitId":"u_165","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_182","unitId":"u_165","date":"2026-04-18","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UDI4S19MUS","remarks":"","createdBy":"Import","createdAt":"2026-04-18"},{"id":"tx_183","unitId":"u_165","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_184","unitId":"u_165","date":"2026-05-06","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1200,"mpesaRef":"UE64S3C2RF","remarks":"","createdBy":"Import","createdAt":"2026-05-06"},{"id":"tx_185","unitId":"u_165","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_186","unitId":"u_165","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_187","unitId":"u_165","date":"2026-07-01","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5200,"mpesaRef":"UG14S9PNPS","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_188","unitId":"u_165","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_189","unitId":"u_165","date":"2026-08-02","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UH24S1F3B7","remarks":"","createdBy":"Import","createdAt":"2026-08-02"},{"id":"tx_192","unitId":"u_190","date":"2024-05-01","month":"2024-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-05-01"},{"id":"tx_193","unitId":"u_190","date":"2024-05-16","month":"2024-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3100,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-05-16"},{"id":"tx_194","unitId":"u_190","date":"2024-06-01","month":"2024-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-06-01"},{"id":"tx_195","unitId":"u_190","date":"2024-06-25","month":"2024-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3100,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-06-25"},{"id":"tx_196","unitId":"u_190","date":"2024-07-01","month":"2024-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-07-01"},{"id":"tx_197","unitId":"u_190","date":"2024-07-18","month":"2024-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3100,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-07-18"},{"id":"tx_198","unitId":"u_190","date":"2024-08-01","month":"2024-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-08-01"},{"id":"tx_199","unitId":"u_190","date":"2024-08-21","month":"2024-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-08-21"},{"id":"tx_200","unitId":"u_190","date":"2024-09-01","month":"2024-09","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-09-01"},{"id":"tx_201","unitId":"u_190","date":"2024-10-01","month":"2024-10","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-10-01"},{"id":"tx_202","unitId":"u_190","date":"2024-10-07","month":"2024-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3100,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-10-07"},{"id":"tx_203","unitId":"u_190","date":"2024-11-01","month":"2024-11","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-11-01"},{"id":"tx_204","unitId":"u_190","date":"2024-11-18","month":"2024-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-11-18"},{"id":"tx_205","unitId":"u_190","date":"2024-12-01","month":"2024-12","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-12-01"},{"id":"tx_206","unitId":"u_190","date":"2024-12-01","month":"2024-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-12-01"},{"id":"tx_207","unitId":"u_190","date":"2025-01-01","month":"2025-01","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-01-01"},{"id":"tx_208","unitId":"u_190","date":"2025-01-04","month":"2025-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"TA48VOSWB2","remarks":"","createdBy":"Import","createdAt":"2025-01-04"},{"id":"tx_209","unitId":"u_190","date":"2025-02-01","month":"2025-02","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-02-01"},{"id":"tx_210","unitId":"u_190","date":"2025-02-17","month":"2025-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TBH1CXME8H","remarks":"","createdBy":"Import","createdAt":"2025-02-17"},{"id":"tx_211","unitId":"u_190","date":"2025-03-01","month":"2025-03","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-03-01"},{"id":"tx_212","unitId":"u_190","date":"2025-03-23","month":"2025-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"TCN2UO3OLA","remarks":"","createdBy":"Import","createdAt":"2025-03-23"},{"id":"tx_213","unitId":"u_190","date":"2025-04-01","month":"2025-04","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-04-01"},{"id":"tx_214","unitId":"u_190","date":"2025-05-01","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-01"},{"id":"tx_215","unitId":"u_190","date":"2025-05-10","month":"2025-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"TEA154FEJT","remarks":"","createdBy":"Import","createdAt":"2025-05-10"},{"id":"tx_216","unitId":"u_190","date":"2025-06-01","month":"2025-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-06-01"},{"id":"tx_217","unitId":"u_190","date":"2025-07-01","month":"2025-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-07-01"},{"id":"tx_218","unitId":"u_190","date":"2025-07-12","month":"2025-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TGC9R2RCKV","remarks":"","createdBy":"Import","createdAt":"2025-07-12"},{"id":"tx_219","unitId":"u_190","date":"2025-08-01","month":"2025-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-08-01"},{"id":"tx_220","unitId":"u_190","date":"2025-08-14","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"THE89DS530","remarks":"","createdBy":"Import","createdAt":"2025-08-14"},{"id":"tx_221","unitId":"u_190","date":"2025-09-01","month":"2025-09","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-09-01"},{"id":"tx_222","unitId":"u_190","date":"2025-09-11","month":"2025-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3100,"mpesaRef":"TIB5D4NWMV","remarks":"","createdBy":"Import","createdAt":"2025-09-11"},{"id":"tx_223","unitId":"u_190","date":"2025-10-01","month":"2025-10","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_224","unitId":"u_190","date":"2025-10-26","month":"2025-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3100,"mpesaRef":"TJQIU8C8Q9","remarks":"","createdBy":"Import","createdAt":"2025-10-26"},{"id":"tx_225","unitId":"u_190","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_226","unitId":"u_190","date":"2025-11-25","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3100,"mpesaRef":"TKPIUB3KFA","remarks":"","createdBy":"Import","createdAt":"2025-11-25"},{"id":"tx_227","unitId":"u_190","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_228","unitId":"u_190","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_229","unitId":"u_190","date":"2026-01-06","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"ua6iu2y99K","remarks":"","createdBy":"Import","createdAt":"2026-01-06"},{"id":"tx_230","unitId":"u_190","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_231","unitId":"u_190","date":"2026-02-20","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4100,"mpesaRef":"UBKIU73PZH","remarks":"","createdBy":"Import","createdAt":"2026-02-20"},{"id":"tx_232","unitId":"u_190","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_233","unitId":"u_190","date":"2026-03-19","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4100,"mpesaRef":"UCJIU9OJZN","remarks":"","createdBy":"Import","createdAt":"2026-03-19"},{"id":"tx_234","unitId":"u_190","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_235","unitId":"u_190","date":"2026-04-25","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"UDPIU20RHL","remarks":"","createdBy":"Import","createdAt":"2026-04-25"},{"id":"tx_236","unitId":"u_190","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_237","unitId":"u_190","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_238","unitId":"u_190","date":"2026-06-10","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"UFAIU76PUE","remarks":"","createdBy":"Import","createdAt":"2026-06-10"},{"id":"tx_239","unitId":"u_190","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_240","unitId":"u_190","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":100,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_241","unitId":"u_190","date":"2026-08-01","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"","remarks":"CASH","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_244","unitId":"u_242","date":"2024-10-01","month":"2024-10","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-10-01"},{"id":"tx_245","unitId":"u_242","date":"2024-10-30","month":"2024-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3500,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-10-30"},{"id":"tx_246","unitId":"u_242","date":"2024-11-01","month":"2024-11","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-11-01"},{"id":"tx_247","unitId":"u_242","date":"2024-12-01","month":"2024-12","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-12-01"},{"id":"tx_248","unitId":"u_242","date":"2024-12-27","month":"2024-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-12-27"},{"id":"tx_249","unitId":"u_242","date":"2025-01-01","month":"2025-01","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-01-01"},{"id":"tx_250","unitId":"u_242","date":"2025-01-01","month":"2025-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-01-01"},{"id":"tx_251","unitId":"u_242","date":"2025-02-01","month":"2025-02","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-02-01"},{"id":"tx_252","unitId":"u_242","date":"2025-02-01","month":"2025-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-02-01"},{"id":"tx_253","unitId":"u_242","date":"2025-03-01","month":"2025-03","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-03-01"},{"id":"tx_254","unitId":"u_242","date":"2025-03-07","month":"2025-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"TC70SR0JO2","remarks":"","createdBy":"Import","createdAt":"2025-03-07"},{"id":"tx_255","unitId":"u_242","date":"2025-03-20","month":"2025-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1000,"mpesaRef":"TCK4FJKOF6","remarks":"","createdBy":"Import","createdAt":"2025-03-20"},{"id":"tx_256","unitId":"u_242","date":"2025-04-01","month":"2025-04","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-04-01"},{"id":"tx_257","unitId":"u_242","date":"2025-05-01","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-01"},{"id":"tx_258","unitId":"u_242","date":"2025-05-16","month":"2025-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":9600,"mpesaRef":"TEG4WN2RS4","remarks":"","createdBy":"Import","createdAt":"2025-05-16"},{"id":"tx_259","unitId":"u_242","date":"2025-06-01","month":"2025-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-06-01"},{"id":"tx_260","unitId":"u_242","date":"2025-06-28","month":"2025-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TFS0RQ4RVA","remarks":"","createdBy":"Import","createdAt":"2025-06-28"},{"id":"tx_261","unitId":"u_242","date":"2025-07-01","month":"2025-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-07-01"},{"id":"tx_262","unitId":"u_242","date":"2025-07-21","month":"2025-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"TGL7Y7OV6P","remarks":"","createdBy":"Import","createdAt":"2025-07-21"},{"id":"tx_263","unitId":"u_242","date":"2025-08-01","month":"2025-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-08-01"},{"id":"tx_264","unitId":"u_242","date":"2025-08-15","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"THF0ELQV9Y","remarks":"","createdBy":"Import","createdAt":"2025-08-15"},{"id":"tx_265","unitId":"u_242","date":"2025-09-01","month":"2025-09","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-09-01"},{"id":"tx_266","unitId":"u_242","date":"2025-10-01","month":"2025-10","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_267","unitId":"u_242","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_268","unitId":"u_242","date":"2025-11-05","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1500,"mpesaRef":"TK5DC9D8XI","remarks":"","createdBy":"Import","createdAt":"2025-11-05"},{"id":"tx_269","unitId":"u_242","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_270","unitId":"u_242","date":"2025-12-05","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3500,"mpesaRef":"TL5DC06X48","remarks":"","createdBy":"Import","createdAt":"2025-12-05"},{"id":"tx_271","unitId":"u_242","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_272","unitId":"u_242","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_273","unitId":"u_242","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_274","unitId":"u_242","date":"2026-03-09","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4600,"mpesaRef":"UC9HU8OTKI","remarks":"","createdBy":"Import","createdAt":"2026-03-09"},{"id":"tx_275","unitId":"u_242","date":"2026-03-10","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":9700,"mpesaRef":"UCAHU8SG1O","remarks":"","createdBy":"Import","createdAt":"2026-03-10"},{"id":"tx_276","unitId":"u_242","date":"2026-03-15","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1500,"mpesaRef":"UCFHU98C55","remarks":"","createdBy":"Import","createdAt":"2026-03-15"},{"id":"tx_277","unitId":"u_242","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_278","unitId":"u_242","date":"2026-04-02","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"UD2HUB3E14","remarks":"","createdBy":"Import","createdAt":"2026-04-02"},{"id":"tx_279","unitId":"u_242","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_280","unitId":"u_242","date":"2026-05-08","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2200,"mpesaRef":"UE8HU3B6A9","remarks":"","createdBy":"Import","createdAt":"2026-05-08"},{"id":"tx_281","unitId":"u_242","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_282","unitId":"u_242","date":"2026-06-26","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2200,"mpesaRef":"UFQHU8V90K","remarks":"","createdBy":"Import","createdAt":"2026-06-26"},{"id":"tx_283","unitId":"u_242","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_284","unitId":"u_242","date":"2026-07-31","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3200,"mpesaRef":"UGVHU0VMOS","remarks":"","createdBy":"Import","createdAt":"2026-07-31"},{"id":"tx_285","unitId":"u_242","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_286","unitId":"u_242","date":"2026-08-30","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4200,"mpesaRef":"UHUHUETOV","remarks":"","createdBy":"Import","createdAt":"2026-08-30"},{"id":"tx_289","unitId":"u_287","date":"2025-09-01","month":"2025-09","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-09-01"},{"id":"tx_290","unitId":"u_287","date":"2025-09-08","month":"2025-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"TI831503H7","remarks":"","createdBy":"Import","createdAt":"2025-09-08"},{"id":"tx_291","unitId":"u_287","date":"2025-10-01","month":"2025-10","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_292","unitId":"u_287","date":"2025-10-08","month":"2025-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"TJ8A66UNIJ","remarks":"","createdBy":"Import","createdAt":"2025-10-08"},{"id":"tx_293","unitId":"u_287","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_294","unitId":"u_287","date":"2025-11-30","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"TK9R59LY8Q","remarks":"","createdBy":"Import","createdAt":"2025-11-30"},{"id":"tx_295","unitId":"u_287","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_296","unitId":"u_287","date":"2025-12-07","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"TL7A6OB9T4","remarks":"","createdBy":"Import","createdAt":"2025-12-07"},{"id":"tx_297","unitId":"u_287","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_298","unitId":"u_287","date":"2026-01-09","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"UA9A638STU","remarks":"","createdBy":"Import","createdAt":"2026-01-09"},{"id":"tx_299","unitId":"u_287","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_300","unitId":"u_287","date":"2026-02-08","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"UB8A661JMU","remarks":"","createdBy":"Import","createdAt":"2026-02-08"},{"id":"tx_301","unitId":"u_287","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_302","unitId":"u_287","date":"2026-03-10","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"UCADZ8VRNX","remarks":"","createdBy":"Import","createdAt":"2026-03-10"},{"id":"tx_303","unitId":"u_287","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_304","unitId":"u_287","date":"2026-04-09","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"UD9DZ085WS","remarks":"","createdBy":"Import","createdAt":"2026-04-09"},{"id":"tx_305","unitId":"u_287","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_306","unitId":"u_287","date":"2026-05-08","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"UE8DZ3MMZ0","remarks":"","createdBy":"Import","createdAt":"2026-05-08"},{"id":"tx_307","unitId":"u_287","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_308","unitId":"u_287","date":"2026-07-08","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"UF8DZ75XWF","remarks":"","createdBy":"Import","createdAt":"2026-07-08"},{"id":"tx_309","unitId":"u_287","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_310","unitId":"u_287","date":"2026-07-07","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"UG7DZAKJEO","remarks":"","createdBy":"Import","createdAt":"2026-07-07"},{"id":"tx_311","unitId":"u_287","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":6000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_312","unitId":"u_287","date":"2026-08-06","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6200,"mpesaRef":"UH6DZ2177G","remarks":"","createdBy":"Import","createdAt":"2026-08-06"},{"id":"tx_315","unitId":"u_313","date":"2024-05-01","month":"2024-05","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-05-01"},{"id":"tx_316","unitId":"u_313","date":"2024-05-07","month":"2024-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6000,"mpesaRef":"SE72GP81IE","remarks":"","createdBy":"Import","createdAt":"2024-05-07"},{"id":"tx_317","unitId":"u_313","date":"2024-06-01","month":"2024-06","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-06-01"},{"id":"tx_318","unitId":"u_313","date":"2024-06-26","month":"2024-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"SFQ7O1P4Q7","remarks":"","createdBy":"Import","createdAt":"2024-06-26"},{"id":"tx_319","unitId":"u_313","date":"2024-07-01","month":"2024-07","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-07-01"},{"id":"tx_320","unitId":"u_313","date":"2024-08-01","month":"2024-08","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-08-01"},{"id":"tx_321","unitId":"u_313","date":"2024-08-05","month":"2024-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":6000,"mpesaRef":"SH54Z1VEDA","remarks":"","createdBy":"Import","createdAt":"2024-08-05"},{"id":"tx_322","unitId":"u_313","date":"2024-08-20","month":"2024-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"SHK8NLDH6I","remarks":"","createdBy":"Import","createdAt":"2024-08-20"},{"id":"tx_323","unitId":"u_313","date":"2024-09-01","month":"2024-09","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-09-01"},{"id":"tx_324","unitId":"u_313","date":"2024-09-15","month":"2024-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"SIFOQ9SMKK","remarks":"","createdBy":"Import","createdAt":"2024-09-15"},{"id":"tx_325","unitId":"u_313","date":"2024-10-01","month":"2024-10","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-10-01"},{"id":"tx_326","unitId":"u_313","date":"2024-10-20","month":"2024-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"SJK4ROWXV2","remarks":"","createdBy":"Import","createdAt":"2024-10-20"},{"id":"tx_327","unitId":"u_313","date":"2024-11-01","month":"2024-11","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-11-01"},{"id":"tx_328","unitId":"u_313","date":"2024-12-01","month":"2024-12","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-12-01"},{"id":"tx_329","unitId":"u_313","date":"2024-12-07","month":"2024-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7000,"mpesaRef":"SL79JCIRMV","remarks":"","createdBy":"Import","createdAt":"2024-12-07"},{"id":"tx_330","unitId":"u_313","date":"2025-01-01","month":"2025-01","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-01-01"},{"id":"tx_331","unitId":"u_313","date":"2025-01-14","month":"2025-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4700,"mpesaRef":"TAE782TMOL","remarks":"","createdBy":"Import","createdAt":"2025-01-14"},{"id":"tx_332","unitId":"u_313","date":"2025-02-01","month":"2025-02","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-02-01"},{"id":"tx_333","unitId":"u_313","date":"2025-02-04","month":"2025-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TB46PKU890","remarks":"","createdBy":"Import","createdAt":"2025-02-04"},{"id":"tx_334","unitId":"u_313","date":"2025-03-01","month":"2025-03","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-03-01"},{"id":"tx_335","unitId":"u_313","date":"2025-03-05","month":"2025-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4700,"mpesaRef":"TC51J9PSCX","remarks":"","createdBy":"Import","createdAt":"2025-03-05"},{"id":"tx_336","unitId":"u_313","date":"2025-04-01","month":"2025-04","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-04-01"},{"id":"tx_337","unitId":"u_313","date":"2025-05-01","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-01"},{"id":"tx_338","unitId":"u_313","date":"2025-05-05","month":"2025-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TE56HZGEJU","remarks":"","createdBy":"Import","createdAt":"2025-05-05"},{"id":"tx_339","unitId":"u_313","date":"2025-06-01","month":"2025-06","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-06-01"},{"id":"tx_340","unitId":"u_313","date":"2025-07-01","month":"2025-07","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-07-01"},{"id":"tx_341","unitId":"u_313","date":"2025-08-01","month":"2025-08","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-08-01"},{"id":"tx_342","unitId":"u_313","date":"2025-08-04","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"TH45XSZUVX","remarks":"","createdBy":"Import","createdAt":"2025-08-04"},{"id":"tx_343","unitId":"u_313","date":"2025-09-01","month":"2025-09","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-09-01"},{"id":"tx_344","unitId":"u_313","date":"2025-10-01","month":"2025-10","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_345","unitId":"u_313","date":"2025-10-22","month":"2025-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"TJMIW828GW","remarks":"","createdBy":"Import","createdAt":"2025-10-22"},{"id":"tx_346","unitId":"u_313","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_347","unitId":"u_313","date":"2025-11-02","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"TK2IW92DMR","remarks":"","createdBy":"Import","createdAt":"2025-11-02"},{"id":"tx_348","unitId":"u_313","date":"2025-11-08","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"TK8IW9NRK8","remarks":"","createdBy":"Import","createdAt":"2025-11-08"},{"id":"tx_349","unitId":"u_313","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_350","unitId":"u_313","date":"2025-12-19","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TLJIW1GYTT","remarks":"","createdBy":"Import","createdAt":"2025-12-19"},{"id":"tx_351","unitId":"u_313","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_352","unitId":"u_313","date":"2026-01-13","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"UADIW3QWRY","remarks":"","createdBy":"Import","createdAt":"2026-01-13"},{"id":"tx_353","unitId":"u_313","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_354","unitId":"u_313","date":"2026-02-04","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"UB4IW5SO5K","remarks":"","createdBy":"Import","createdAt":"2026-02-04"},{"id":"tx_355","unitId":"u_313","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_356","unitId":"u_313","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_357","unitId":"u_313","date":"2026-04-09","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"UD9IW0GL80","remarks":"","createdBy":"Import","createdAt":"2026-04-09"},{"id":"tx_358","unitId":"u_313","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_359","unitId":"u_313","date":"2026-05-04","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":24000,"mpesaRef":"UE4IW3BS4F","remarks":"","createdBy":"Import","createdAt":"2026-05-04"},{"id":"tx_360","unitId":"u_313","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_361","unitId":"u_313","date":"2026-06-08","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":15000,"mpesaRef":"UF8IW7DZAX","remarks":"","createdBy":"Import","createdAt":"2026-06-08"},{"id":"tx_362","unitId":"u_313","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_363","unitId":"u_313","date":"2026-07-08","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4700,"mpesaRef":"UG8IWAXDRB","remarks":"","createdBy":"Import","createdAt":"2026-07-08"},{"id":"tx_364","unitId":"u_313","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_365","unitId":"u_313","date":"2026-08-19","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4700,"mpesaRef":"UHJIW3UZ24","remarks":"","createdBy":"Import","createdAt":"2026-08-19"},{"id":"tx_368","unitId":"u_366","date":"2024-04-01","month":"2024-04","type":"charge","description":"Monthly Rent Charge","rentCharged":30150,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"Balance brought forward","createdBy":"Import","createdAt":"2024-04-01"},{"id":"tx_369","unitId":"u_366","date":"2024-05-01","month":"2024-05","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-05-01"},{"id":"tx_370","unitId":"u_366","date":"2024-05-19","month":"2024-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"SEJ1RV6JUN","remarks":"","createdBy":"Import","createdAt":"2024-05-19"},{"id":"tx_371","unitId":"u_366","date":"2024-06-01","month":"2024-06","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-06-01"},{"id":"tx_372","unitId":"u_366","date":"2024-06-05","month":"2024-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"SF53K4CUNT","remarks":"","createdBy":"Import","createdAt":"2024-06-05"},{"id":"tx_373","unitId":"u_366","date":"2024-07-01","month":"2024-07","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-07-01"},{"id":"tx_374","unitId":"u_366","date":"2024-07-06","month":"2024-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"SG68QT8WCG","remarks":"","createdBy":"Import","createdAt":"2024-07-06"},{"id":"tx_375","unitId":"u_366","date":"2024-08-01","month":"2024-08","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-08-01"},{"id":"tx_376","unitId":"u_366","date":"2024-08-04","month":"2024-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"SH46VEAV96","remarks":"","createdBy":"Import","createdAt":"2024-08-04"},{"id":"tx_377","unitId":"u_366","date":"2024-09-01","month":"2024-09","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-09-01"},{"id":"tx_378","unitId":"u_366","date":"2024-09-06","month":"2024-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"SI68L9BFHK","remarks":"","createdBy":"Import","createdAt":"2024-09-06"},{"id":"tx_379","unitId":"u_366","date":"2024-10-01","month":"2024-10","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-10-01"},{"id":"tx_380","unitId":"u_366","date":"2024-11-01","month":"2024-11","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-11-01"},{"id":"tx_381","unitId":"u_366","date":"2024-12-01","month":"2024-12","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-12-01"},{"id":"tx_382","unitId":"u_366","date":"2024-12-10","month":"2024-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":20000,"mpesaRef":"SLA5ZF5HM1","remarks":"","createdBy":"Import","createdAt":"2024-12-10"},{"id":"tx_383","unitId":"u_366","date":"2025-01-01","month":"2025-01","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-01-01"},{"id":"tx_384","unitId":"u_366","date":"2025-01-15","month":"2025-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"TAF9BPJUP9","remarks":"","createdBy":"Import","createdAt":"2025-01-15"},{"id":"tx_385","unitId":"u_366","date":"2025-02-01","month":"2025-02","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-02-01"},{"id":"tx_386","unitId":"u_366","date":"2025-03-01","month":"2025-03","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-03-01"},{"id":"tx_387","unitId":"u_366","date":"2025-04-01","month":"2025-04","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-04-01"},{"id":"tx_388","unitId":"u_366","date":"2025-04-13","month":"2025-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":9500,"mpesaRef":"TDD0NARXMU","remarks":"","createdBy":"Import","createdAt":"2025-04-13"},{"id":"tx_389","unitId":"u_366","date":"2025-04-15","month":"2025-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TDF9SIFXT1","remarks":"","createdBy":"Import","createdAt":"2025-04-15"},{"id":"tx_390","unitId":"u_366","date":"2025-04-16","month":"2025-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TDG21HOGL8","remarks":"","createdBy":"Import","createdAt":"2025-04-16"},{"id":"tx_391","unitId":"u_366","date":"2025-05-01","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-01"},{"id":"tx_392","unitId":"u_366","date":"2025-05-31","month":"2025-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":20000,"mpesaRef":"TEV7ZFOQYD","remarks":"","createdBy":"Import","createdAt":"2025-05-31"},{"id":"tx_393","unitId":"u_366","date":"2025-06-01","month":"2025-06","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-06-01"},{"id":"tx_394","unitId":"u_366","date":"2025-07-01","month":"2025-07","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-07-01"},{"id":"tx_395","unitId":"u_366","date":"2025-08-01","month":"2025-08","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-08-01"},{"id":"tx_396","unitId":"u_366","date":"2025-08-11","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"THB1XMCKO5","remarks":"","createdBy":"Import","createdAt":"2025-08-11"},{"id":"tx_397","unitId":"u_366","date":"2025-08-22","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"THM8IGI8CQ","remarks":"","createdBy":"Import","createdAt":"2025-08-22"},{"id":"tx_398","unitId":"u_366","date":"2025-08-31","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"THV9SZVOK1","remarks":"","createdBy":"Import","createdAt":"2025-08-31"},{"id":"tx_399","unitId":"u_366","date":"2025-09-01","month":"2025-09","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-09-01"},{"id":"tx_400","unitId":"u_366","date":"2025-10-01","month":"2025-10","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_401","unitId":"u_366","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_402","unitId":"u_366","date":"2025-11-02","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4350,"mpesaRef":"TK29R92G87","remarks":"","createdBy":"Import","createdAt":"2025-11-02"},{"id":"tx_403","unitId":"u_366","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_404","unitId":"u_366","date":"2025-12-14","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"TLE9R10L9C","remarks":"","createdBy":"Import","createdAt":"2025-12-14"},{"id":"tx_405","unitId":"u_366","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_406","unitId":"u_366","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_407","unitId":"u_366","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_408","unitId":"u_366","date":"2026-03-01","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":22100,"mpesaRef":"UC19R84RTZ","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_409","unitId":"u_366","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_410","unitId":"u_366","date":"2026-04-04","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"UD49RBOFZ1","remarks":"","createdBy":"Import","createdAt":"2026-04-04"},{"id":"tx_411","unitId":"u_366","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_412","unitId":"u_366","date":"2026-05-13","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":22400,"mpesaRef":"UED9R43L46","remarks":"","createdBy":"Import","createdAt":"2026-05-13"},{"id":"tx_413","unitId":"u_366","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_414","unitId":"u_366","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_415","unitId":"u_366","date":"2026-07-01","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"UG19R9RBDD4","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_416","unitId":"u_366","date":"2026-07-26","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":7700,"mpesaRef":"UGQ9R0O6HD","remarks":"","createdBy":"Import","createdAt":"2026-07-26"},{"id":"tx_417","unitId":"u_366","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":7500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_418","unitId":"u_366","date":"2026-08-10","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":8100,"mpesaRef":"UHA9R2FE4U","remarks":"","createdBy":"Import","createdAt":"2026-08-10"},{"id":"tx_419","unitId":"u_366","date":"2026-08-25","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":15000,"mpesaRef":"UHP9R41XGC","remarks":"","createdBy":"Import","createdAt":"2026-08-25"},{"id":"tx_422","unitId":"u_420","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":4500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"1,000 Paid for Deposit","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_423","unitId":"u_420","date":"2026-03-09","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4500,"mpesaRef":"UC9O28Q3OS","remarks":"","createdBy":"Import","createdAt":"2026-03-09"},{"id":"tx_424","unitId":"u_420","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":3500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_425","unitId":"u_420","date":"2026-04-14","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"UDEGU1424O","remarks":"SAMUEL","createdBy":"Import","createdAt":"2026-04-14"},{"id":"tx_426","unitId":"u_420","date":"2026-04-20","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1000,"mpesaRef":"UDKGU1S6U5","remarks":"SAMUEL","createdBy":"Import","createdAt":"2026-04-20"},{"id":"tx_427","unitId":"u_420","date":"2026-04-20","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":500,"mpesaRef":"UDKGU1T31J","remarks":"SAMUEL","createdBy":"Import","createdAt":"2026-04-20"},{"id":"tx_428","unitId":"u_420","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_429","unitId":"u_420","date":"2026-05-25","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"UEPQR5EBFV","remarks":"ARON","createdBy":"Import","createdAt":"2026-05-25"},{"id":"tx_430","unitId":"u_420","date":"2026-05-26","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1500,"mpesaRef":"UEQGU5X00A","remarks":"SAMUEL","createdBy":"Import","createdAt":"2026-05-26"},{"id":"tx_431","unitId":"u_420","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_432","unitId":"u_420","date":"2026-06-13","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1850,"mpesaRef":"UFDO37LF3A","remarks":"DEDAN","createdBy":"Import","createdAt":"2026-06-13"},{"id":"tx_433","unitId":"u_420","date":"2026-06-22","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1850,"mpesaRef":"UFMQR8MJPF","remarks":"ARON","createdBy":"Import","createdAt":"2026-06-22"},{"id":"tx_434","unitId":"u_420","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_435","unitId":"u_420","date":"2026-07-09","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1850,"mpesaRef":"UG9O2AHJCZ","remarks":"DEDAN","createdBy":"Import","createdAt":"2026-07-09"},{"id":"tx_436","unitId":"u_420","date":"2026-07-18","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1850,"mpesaRef":"UGIQRBLQJ0","remarks":"ARON","createdBy":"Import","createdAt":"2026-07-18"},{"id":"tx_437","unitId":"u_420","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":3000,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_438","unitId":"u_420","date":"2026-08-09","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1850,"mpesaRef":"UH9QR25U7I","remarks":"ARON Morompi","createdBy":"Import","createdAt":"2026-08-09"},{"id":"tx_439","unitId":"u_420","date":"2026-08-10","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":450,"mpesaRef":"UHAUH2KSLI","remarks":"DEDAN","createdBy":"Import","createdAt":"2026-08-10"},{"id":"tx_442","unitId":"u_440","date":"2024-05-01","month":"2024-05","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-05-01"},{"id":"tx_443","unitId":"u_440","date":"2024-05-10","month":"2024-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"SEA251W1AK","remarks":"","createdBy":"Import","createdAt":"2024-05-10"},{"id":"tx_444","unitId":"u_440","date":"2024-05-10","month":"2024-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":200,"mpesaRef":"SEA6S22XEC","remarks":"","createdBy":"Import","createdAt":"2024-05-10"},{"id":"tx_445","unitId":"u_440","date":"2024-05-26","month":"2024-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1500,"mpesaRef":"SEQ7HDVU8D","remarks":"","createdBy":"Import","createdAt":"2024-05-26"},{"id":"tx_446","unitId":"u_440","date":"2024-06-06","month":"2024-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"SF66KH05E0","remarks":"1,700 carried forward from here","createdBy":"Import","createdAt":"2024-06-06"},{"id":"tx_447","unitId":"u_440","date":"2024-06-30","month":"2024-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1700,"mpesaRef":"","remarks":"Cleared","createdBy":"Import","createdAt":"2024-06-30"},{"id":"tx_448","unitId":"u_440","date":"2024-07-01","month":"2024-07","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-07-01"},{"id":"tx_449","unitId":"u_440","date":"2024-07-09","month":"2024-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"SG945RONMS","remarks":"","createdBy":"Import","createdAt":"2024-07-09"},{"id":"tx_450","unitId":"u_440","date":"2024-07-14","month":"2024-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1700,"mpesaRef":"SGE3ODZ47Z","remarks":"","createdBy":"Import","createdAt":"2024-07-14"},{"id":"tx_451","unitId":"u_440","date":"2025-08-01","month":"2025-08","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-08-01"},{"id":"tx_452","unitId":"u_440","date":"2024-08-10","month":"2024-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5700,"mpesaRef":"SHA7KCOGKJ","remarks":"","createdBy":"Import","createdAt":"2024-08-10"},{"id":"tx_453","unitId":"u_440","date":"2024-09-01","month":"2024-09","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-09-01"},{"id":"tx_454","unitId":"u_440","date":"2025-09-10","month":"2025-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5700,"mpesaRef":"SIA65XIRXU","remarks":"","createdBy":"Import","createdAt":"2025-09-10"},{"id":"tx_455","unitId":"u_440","date":"2024-10-01","month":"2024-10","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-10-01"},{"id":"tx_456","unitId":"u_440","date":"2024-10-07","month":"2024-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5700,"mpesaRef":"SJ778ERW8X","remarks":"","createdBy":"Import","createdAt":"2024-10-07"},{"id":"tx_457","unitId":"u_440","date":"2024-11-01","month":"2024-11","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-11-01"},{"id":"tx_458","unitId":"u_440","date":"2024-11-10","month":"2024-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5700,"mpesaRef":"SKA6C344KI","remarks":"","createdBy":"Import","createdAt":"2024-11-10"},{"id":"tx_459","unitId":"u_440","date":"2024-12-01","month":"2024-12","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2024-12-01"},{"id":"tx_460","unitId":"u_440","date":"2024-12-10","month":"2024-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5700,"mpesaRef":"SLA3XGG2CJ","remarks":"","createdBy":"Import","createdAt":"2024-12-10"},{"id":"tx_461","unitId":"u_440","date":"2025-01-01","month":"2025-01","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-01-01"},{"id":"tx_462","unitId":"u_440","date":"2025-01-07","month":"2025-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5700,"mpesaRef":"TA74AORM3M","remarks":"","createdBy":"Import","createdAt":"2025-01-07"},{"id":"tx_463","unitId":"u_440","date":"2025-02-01","month":"2025-02","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-02-01"},{"id":"tx_464","unitId":"u_440","date":"2025-02-10","month":"2025-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5700,"mpesaRef":"TBA7H9QD1L","remarks":"","createdBy":"Import","createdAt":"2025-02-10"},{"id":"tx_465","unitId":"u_440","date":"2025-03-01","month":"2025-03","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-03-01"},{"id":"tx_466","unitId":"u_440","date":"2025-03-09","month":"2025-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TC96ZANLGG","remarks":"","createdBy":"Import","createdAt":"2025-03-09"},{"id":"tx_467","unitId":"u_440","date":"2025-03-14","month":"2025-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"TCE5OBYOCP","remarks":"","createdBy":"Import","createdAt":"2025-03-14"},{"id":"tx_468","unitId":"u_440","date":"2025-04-01","month":"2025-04","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-04-01"},{"id":"tx_469","unitId":"u_440","date":"2025-04-10","month":"2025-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TDA59CU3FB","remarks":"","createdBy":"Import","createdAt":"2025-04-10"},{"id":"tx_470","unitId":"u_440","date":"2025-04-15","month":"2025-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"TDF6SGGA1K","remarks":"","createdBy":"Import","createdAt":"2025-04-15"},{"id":"tx_471","unitId":"u_440","date":"2025-05-01","month":"2025-05","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-05-01"},{"id":"tx_472","unitId":"u_440","date":"2025-05-11","month":"2025-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TEB0CKS3C0","remarks":"","createdBy":"Import","createdAt":"2025-05-11"},{"id":"tx_473","unitId":"u_440","date":"2025-05-17","month":"2025-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"TEH72BUUQX","remarks":"","createdBy":"Import","createdAt":"2025-05-17"},{"id":"tx_474","unitId":"u_440","date":"2025-06-01","month":"2025-06","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-06-01"},{"id":"tx_475","unitId":"u_440","date":"2025-06-11","month":"2025-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TFB7E1F15L","remarks":"","createdBy":"Import","createdAt":"2025-06-11"},{"id":"tx_476","unitId":"u_440","date":"2025-06-27","month":"2025-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"TFR5J6VRYD","remarks":"","createdBy":"Import","createdAt":"2025-06-27"},{"id":"tx_477","unitId":"u_440","date":"2025-07-01","month":"2025-07","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-07-01"},{"id":"tx_478","unitId":"u_440","date":"2025-07-11","month":"2025-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"TGB5HP72VN","remarks":"","createdBy":"Import","createdAt":"2025-07-11"},{"id":"tx_479","unitId":"u_440","date":"2025-07-17","month":"2025-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"TGH8F0MEES","remarks":"","createdBy":"Import","createdAt":"2025-07-17"},{"id":"tx_480","unitId":"u_440","date":"2025-07-28","month":"2025-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"TGS5X49E4V","remarks":"","createdBy":"Import","createdAt":"2025-07-28"},{"id":"tx_481","unitId":"u_440","date":"2025-08-11","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"THB3UZO1CV","remarks":"","createdBy":"Import","createdAt":"2025-08-11"},{"id":"tx_482","unitId":"u_440","date":"2025-08-12","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"THC43SKR88","remarks":"","createdBy":"Import","createdAt":"2025-08-12"},{"id":"tx_483","unitId":"u_440","date":"2025-09-01","month":"2025-09","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-09-01"},{"id":"tx_484","unitId":"u_440","date":"2025-09-10","month":"2025-09","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5700,"mpesaRef":"TIA28JWG7U","remarks":"","createdBy":"Import","createdAt":"2025-09-10"},{"id":"tx_485","unitId":"u_440","date":"2025-10-01","month":"2025-10","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_486","unitId":"u_440","date":"2025-10-09","month":"2025-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TJ9D76ZIUI","remarks":"","createdBy":"Import","createdAt":"2025-10-09"},{"id":"tx_487","unitId":"u_440","date":"2025-10-12","month":"2025-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"TJCD7782FX","remarks":"","createdBy":"Import","createdAt":"2025-10-12"},{"id":"tx_488","unitId":"u_440","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_489","unitId":"u_440","date":"2025-11-10","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5400,"mpesaRef":"TKAD79W04K","remarks":"","createdBy":"Import","createdAt":"2025-11-10"},{"id":"tx_490","unitId":"u_440","date":"2025-11-15","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":300,"mpesaRef":"TKFD7ACEI5","remarks":"","createdBy":"Import","createdAt":"2025-11-15"},{"id":"tx_491","unitId":"u_440","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_492","unitId":"u_440","date":"2025-12-10","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"TLAD70PGME","remarks":"","createdBy":"Import","createdAt":"2025-12-10"},{"id":"tx_493","unitId":"u_440","date":"2025-12-13","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"TLDD70Z86L","remarks":"","createdBy":"Import","createdAt":"2025-12-13"},{"id":"tx_494","unitId":"u_440","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_495","unitId":"u_440","date":"2026-01-10","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"UAAD73MF1N","remarks":"","createdBy":"Import","createdAt":"2026-01-10"},{"id":"tx_496","unitId":"u_440","date":"2026-01-29","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"UATD75CB7K","remarks":"","createdBy":"Import","createdAt":"2026-01-29"},{"id":"tx_497","unitId":"u_440","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_498","unitId":"u_440","date":"2026-02-11","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2700,"mpesaRef":"UBBD76MB0Q","remarks":"","createdBy":"Import","createdAt":"2026-02-11"},{"id":"tx_499","unitId":"u_440","date":"2026-02-20","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3000,"mpesaRef":"UBKD77IF6B","remarks":"","createdBy":"Import","createdAt":"2026-02-20"},{"id":"tx_500","unitId":"u_440","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_501","unitId":"u_440","date":"2026-03-10","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"UCAD798LB7","remarks":"","createdBy":"Import","createdAt":"2026-03-10"},{"id":"tx_502","unitId":"u_440","date":"2026-03-30","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"UCUD7BBSS4","remarks":"","createdBy":"Import","createdAt":"2026-03-30"},{"id":"tx_503","unitId":"u_440","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_504","unitId":"u_440","date":"2026-04-10","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"UDAD70QPBY","remarks":"","createdBy":"Import","createdAt":"2026-04-10"},{"id":"tx_505","unitId":"u_440","date":"2026-04-24","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"UDOD72C3UD","remarks":"","createdBy":"Import","createdAt":"2026-04-24"},{"id":"tx_506","unitId":"u_440","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_507","unitId":"u_440","date":"2026-05-10","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":2000,"mpesaRef":"UEAD7474Y9","remarks":"","createdBy":"Import","createdAt":"2026-05-10"},{"id":"tx_508","unitId":"u_440","date":"2026-05-27","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":3700,"mpesaRef":"UEPD75W5NC","remarks":"","createdBy":"Import","createdAt":"2026-05-27"},{"id":"tx_509","unitId":"u_440","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_510","unitId":"u_440","date":"2026-06-10","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":5000,"mpesaRef":"UFAD77PMMY","remarks":"","createdBy":"Import","createdAt":"2026-06-10"},{"id":"tx_511","unitId":"u_440","date":"2026-06-22","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":700,"mpesaRef":"UFMD797GYN","remarks":"","createdBy":"Import","createdAt":"2026-06-22"},{"id":"tx_512","unitId":"u_440","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_513","unitId":"u_440","date":"2026-07-10","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4000,"mpesaRef":"UGAD7BBHHE","remarks":"","createdBy":"Import","createdAt":"2026-07-10"},{"id":"tx_514","unitId":"u_440","date":"2026-07-16","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":1700,"mpesaRef":"UGGD703GR1","remarks":"","createdBy":"Import","createdAt":"2026-07-16"},{"id":"tx_515","unitId":"u_440","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":5500,"garbageCharged":200,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_516","unitId":"u_440","date":"2026-08-10","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":4400,"mpesaRef":"UHAD72ZSSD","remarks":"","createdBy":"Import","createdAt":"2026-08-10"},{"id":"tx_519","unitId":"u_517","date":"2025-08-01","month":"2025-08","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"10,000 for Deposit (Never Paid)","createdBy":"Import","createdAt":"2025-08-01"},{"id":"tx_520","unitId":"u_517","date":"2025-08-29","month":"2025-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"TNT0I9IQXQ","remarks":"Paid for September","createdBy":"Import","createdAt":"2025-08-29"},{"id":"tx_521","unitId":"u_517","date":"2025-10-01","month":"2025-10","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_522","unitId":"u_517","date":"2025-10-03","month":"2025-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"TJ33O6ES5V","remarks":"","createdBy":"Import","createdAt":"2025-10-03"},{"id":"tx_523","unitId":"u_517","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_524","unitId":"u_517","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_525","unitId":"u_517","date":"2025-12-23","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"TLN3O220O89","remarks":"","createdBy":"Import","createdAt":"2025-12-23"},{"id":"tx_526","unitId":"u_517","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_527","unitId":"u_517","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_528","unitId":"u_517","date":"2026-02-12","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":20000,"mpesaRef":"UBC306X9AG","remarks":"","createdBy":"Import","createdAt":"2026-02-12"},{"id":"tx_529","unitId":"u_517","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_530","unitId":"u_517","date":"2026-03-10","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UCA3O9FAR9","remarks":"","createdBy":"Import","createdAt":"2026-03-10"},{"id":"tx_531","unitId":"u_517","date":"2026-03-16","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UCG3O9ZK6A","remarks":"","createdBy":"Import","createdAt":"2026-03-16"},{"id":"tx_532","unitId":"u_517","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_533","unitId":"u_517","date":"2026-04-17","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UDH3O1QHCU","remarks":"","createdBy":"Import","createdAt":"2026-04-17"},{"id":"tx_534","unitId":"u_517","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_535","unitId":"u_517","date":"2026-05-12","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UECOM3XS7A","remarks":"","createdBy":"Import","createdAt":"2026-05-12"},{"id":"tx_536","unitId":"u_517","date":"2026-06-01","month":"2026-06","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-06-01"},{"id":"tx_537","unitId":"u_517","date":"2026-06-15","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UFFLO7UNIQ","remarks":"","createdBy":"Import","createdAt":"2026-06-15"},{"id":"tx_538","unitId":"u_517","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_539","unitId":"u_517","date":"2026-07-13","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"","remarks":"CASH","createdBy":"Import","createdAt":"2026-07-13"},{"id":"tx_540","unitId":"u_517","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_541","unitId":"u_517","date":"2026-08-10","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UHALO28D5G","remarks":"","createdBy":"Import","createdAt":"2026-08-10"},{"id":"tx_544","unitId":"u_542","date":"2025-10-01","month":"2025-10","type":"charge","description":"Monthly Rent Charge","rentCharged":15000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"5,000 For Deposit","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_545","unitId":"u_542","date":"2025-10-01","month":"2025-10","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":15000,"mpesaRef":"TJ1QX67JDX","remarks":"","createdBy":"Import","createdAt":"2025-10-01"},{"id":"tx_546","unitId":"u_542","date":"2025-11-01","month":"2025-11","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-11-01"},{"id":"tx_547","unitId":"u_542","date":"2025-11-03","month":"2025-11","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"TK3QX965S0","remarks":"","createdBy":"Import","createdAt":"2025-11-03"},{"id":"tx_548","unitId":"u_542","date":"2025-12-01","month":"2025-12","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2025-12-01"},{"id":"tx_549","unitId":"u_542","date":"2025-12-04","month":"2025-12","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"TL4QX04BIY","remarks":"","createdBy":"Import","createdAt":"2025-12-04"},{"id":"tx_550","unitId":"u_542","date":"2026-01-01","month":"2026-01","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-01-01"},{"id":"tx_551","unitId":"u_542","date":"2026-01-05","month":"2026-01","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UA57G2WSPL","remarks":"","createdBy":"Import","createdAt":"2026-01-05"},{"id":"tx_552","unitId":"u_542","date":"2026-02-01","month":"2026-02","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-02-01"},{"id":"tx_553","unitId":"u_542","date":"2026-02-05","month":"2026-02","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UB5QX5XU12","remarks":"","createdBy":"Import","createdAt":"2026-02-05"},{"id":"tx_554","unitId":"u_542","date":"2026-03-01","month":"2026-03","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-03-01"},{"id":"tx_555","unitId":"u_542","date":"2026-03-06","month":"2026-03","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UC67G8GXVD","remarks":"","createdBy":"Import","createdAt":"2026-03-06"},{"id":"tx_556","unitId":"u_542","date":"2026-04-01","month":"2026-04","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-04-01"},{"id":"tx_557","unitId":"u_542","date":"2026-04-06","month":"2026-04","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UD67GBSFFK","remarks":"","createdBy":"Import","createdAt":"2026-04-06"},{"id":"tx_558","unitId":"u_542","date":"2026-05-01","month":"2026-05","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-05-01"},{"id":"tx_559","unitId":"u_542","date":"2026-05-05","month":"2026-05","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UE57G390K3","remarks":"","createdBy":"Import","createdAt":"2026-05-05"},{"id":"tx_560","unitId":"u_542","date":"2026-06-06","month":"2026-06","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UF6QX73FMK","remarks":"","createdBy":"Import","createdAt":"2026-06-06"},{"id":"tx_561","unitId":"u_542","date":"2026-07-01","month":"2026-07","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-07-01"},{"id":"tx_562","unitId":"u_542","date":"2026-07-05","month":"2026-07","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UG57GA8J6G","remarks":"","createdBy":"Import","createdAt":"2026-07-05"},{"id":"tx_563","unitId":"u_542","date":"2026-08-01","month":"2026-08","type":"charge","description":"Monthly Rent Charge","rentCharged":10000,"garbageCharged":0,"paymentReceived":0,"mpesaRef":"","remarks":"","createdBy":"Import","createdAt":"2026-08-01"},{"id":"tx_564","unitId":"u_542","date":"2026-08-05","month":"2026-08","type":"payment","description":"Payment Received","rentCharged":0,"garbageCharged":0,"paymentReceived":10000,"mpesaRef":"UH57G1TVAH","remarks":"","createdBy":"Import","createdAt":"2026-08-05"}],"archive":[]};
}

/* ---------------------------------------------------------------------- */
/* Ledger math (Approach B: derived balances, never stored)                */
/* ---------------------------------------------------------------------- */
function ledgerWithRunningBalance(transactions) {
  const sorted = [...transactions].sort((a, b) => (String(a.date || "") + a.id).localeCompare(String(b.date || "") + b.id));
  let bal = 0;
  return sorted.map((t) => {
    bal += (t.rentCharged || 0) + (t.garbageCharged || 0) - (t.paymentReceived || 0);
    return { ...t, runningBalance: bal };
  });
}
function unitBalance(transactions) {
  const ledger = ledgerWithRunningBalance(transactions);
  return ledger.length ? ledger[ledger.length - 1].runningBalance : 0;
}
function monthsOutstanding(balance, rent, garbage) {
  const denom = (rent || 0) + (garbage || 0);
  if (balance <= 0 || denom <= 0) return 0;
  return balance / denom;
}
function statusOf(unit, balance) {
  if (unit.occupancyStatus === "vacant") return "Vacant";
  if (balance <= 0) return "Paid/Credit";
  const mo = monthsOutstanding(balance, unit.rent, unit.garbage);
  if (mo < 1) return "<1 month";
  if (mo < 2) return "1–2 months";
  if (mo < 3) return "2–3 months";
  return ">3 months";
}
function priorityOf(status) {
  switch (status) {
    case "Vacant": return { key: "VACANT", action: "Find tenant" };
    case "Paid/Credit": return { key: "PAID", action: "No action needed" };
    case "<1 month": return { key: "LOW", action: "Monitor" };
    case "1–2 months": return { key: "MEDIUM", action: "Follow up" };
    default: return { key: "HIGH", action: "Urgent" };
  }
}
function greeting() {
  const h = new Date().getHours();
  if (h < 12) return "Good morning";
  if (h < 17) return "Good afternoon";
  return "Good evening";
}

// Drives the sidebar/mobile-nav split via JS rather than a CSS breakpoint class,
// since this environment only ships a precompiled subset of Tailwind utilities
// and responsive variants (e.g. "sm:flex") are not guaranteed to be included.
function useIsDesktop(breakpoint = 640) {
  const [isDesktop, setIsDesktop] = useState(() => {
    try { return window.innerWidth >= breakpoint; } catch { return true; }
  });
  useEffect(() => {
    let mq;
    const handleResize = () => setIsDesktop(window.innerWidth >= breakpoint);
    try {
      mq = window.matchMedia(`(min-width: ${breakpoint}px)`);
      const onChange = (e) => setIsDesktop(e.matches);
      mq.addEventListener ? mq.addEventListener("change", onChange) : mq.addListener(onChange);
      setIsDesktop(mq.matches);
      return () => { mq.removeEventListener ? mq.removeEventListener("change", onChange) : mq.removeListener(onChange); };
    } catch {
      window.addEventListener("resize", handleResize);
      return () => window.removeEventListener("resize", handleResize);
    }
  }, [breakpoint]);
  return isDesktop;
}

/* ---------------------------------------------------------------------- */
/* Small UI atoms                                                          */
/* ---------------------------------------------------------------------- */
function Pill({ label, colorKey }) {
  const c = STATUS_COLORS[colorKey] || STATUS_COLORS.VACANT;
  return (
    <span
      className="inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold whitespace-nowrap"
      style={{ background: c.bg, color: c.fg }}
    >
      <span className="w-1.5 h-1.5 rounded-full" style={{ background: c.dot }} />
      {label}
    </span>
  );
}

function Btn({ children, onClick, variant = "primary", disabled, size = "md", icon: Icon, type = "button" }) {
  const base = "inline-flex items-center justify-center gap-1.5 rounded-xl font-semibold transition-colors disabled:opacity-40 disabled:cursor-not-allowed";
  const sizes = { sm: "px-2.5 py-1.5 text-xs", md: "px-4 py-2.5 text-sm" };
  const variants = {
    primary: "text-white",
    secondary: "bg-white border hover:bg-black/[0.02]",
    ghost: "bg-transparent hover:bg-black/5",
    danger: "bg-red-50 text-red-700 hover:bg-red-100",
  };
  const style =
    variant === "primary" ? { background: "#1F8A4C" } :
    variant === "secondary" ? { borderColor: BORDER, color: INK } : {};
  return (
    <button type={type} disabled={disabled} onClick={onClick} className={`${base} ${sizes[size]} ${variants[variant]}`} style={style}>
      {Icon && <Icon size={14} />}
      {children}
    </button>
  );
}

// Compact per-row "..." menu used for Edit/Delete on individual ledger
// transactions — keeps the ledger table from getting cluttered with buttons.
function RowActionsMenu({ onEdit, onDelete, editLabel = "Edit", deleteLabel = "Delete" }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const onDocClick = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, [open]);
  return (
    <div className="relative inline-block" ref={ref}>
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        className="p-1.5 rounded-lg hover:bg-black/5"
        aria-label="Transaction actions"
      >
        <MoreVertical size={16} style={{ color: MUTED }} />
      </button>
      {open && (
        <div className="absolute right-0 z-20 mt-1 w-40 rounded-xl border bg-white shadow-lg overflow-hidden" style={{ borderColor: BORDER }}>
          <button
            type="button"
            onClick={() => { setOpen(false); onEdit(); }}
            className="w-full flex items-center gap-2 px-3 py-2 text-xs font-semibold text-left hover:bg-black/[0.03]"
          >
            <Pencil size={13} /> {editLabel}
          </button>
          <button
            type="button"
            onClick={() => { setOpen(false); onDelete(); }}
            className="w-full flex items-center gap-2 px-3 py-2 text-xs font-semibold text-left hover:bg-red-50"
            style={{ color: "#C0392B" }}
          >
            <Trash2 size={13} /> {deleteLabel}
          </button>
        </div>
      )}
    </div>
  );
}

function Field({ label, children, hint }) {
  return (
    <label className="block mb-3">
      <div className="text-xs font-semibold mb-1" style={{ color: MUTED }}>{label}</div>
      {children}
      {hint && <div className="text-[11px] mt-1" style={{ color: MUTED }}>{hint}</div>}
    </label>
  );
}
const inputCls = "w-full rounded-xl border px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-emerald-100";
const inputStyle = { borderColor: BORDER };

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-black/40 p-0 sm:p-4">
      <div className={`bg-white w-full ${wide ? "sm:max-w-2xl" : "sm:max-w-md"} sm:rounded-2xl rounded-t-2xl max-h-[92vh] overflow-y-auto`}>
        <div className="flex items-center justify-between px-5 py-4 border-b sticky top-0 bg-white z-10" style={{ borderColor: BORDER }}>
          <h3 className="font-bold text-base" style={{ color: INK }}>{title}</h3>
          <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-black/5"><X size={18} /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Card({ children, tone, className = "" }) {
  const toneBg = tone === "green" ? "#EDF8F1" : tone === "red" ? "#FCEBEB" : tone === "amber" ? "#FDF3E2" : CARD;
  return (
    <div className={`rounded-2xl border p-4 ${className}`} style={{ background: toneBg, borderColor: BORDER }}>
      {children}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Sidebar                                                                  */
/* ---------------------------------------------------------------------- */
const NAV_ITEMS = [
  { key: "overview", label: "Overview", icon: Home },
  { key: "units", label: "Current units", icon: ListChecks },
  { key: "archive", label: "Closed tenancies", icon: ArchiveIcon },
  { key: "importexport", label: "Import & export", icon: UploadCloud },
  { key: "users", label: "Users & access", icon: Users, adminOnly: true },
];

function Sidebar({ screen, go, role, onLogout }) {
  const initials = role === "Owner/Admin" ? "OA" : role === "Manager" ? "MG" : "RO";
  return (
    <aside
      className="flex flex-col shrink-0 sticky top-0 overflow-y-auto"
      style={{ background: SIDEBAR, width: 260, height: "100vh" }}
    >
      <div className="px-5 pt-6 pb-5">
        <button
          type="button"
          onClick={() => go("overview")}
          aria-label="Go to MaliDesk Home dashboard"
          className="w-full text-left rounded-2xl p-2 -m-2 transition-opacity hover:opacity-90 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-offset-[#1A3C34]"
          style={{ focusRingColor: AMBER }}
        >
          <div className="flex items-center gap-3">
            <div
              className="w-11 h-11 rounded-xl flex items-center justify-center shrink-0"
              style={{ background: AMBER, color: SIDEBAR }}
            >
              <svg viewBox="0 0 64 64" width="27" height="27" aria-hidden="true">
                <path d="M16 18h9v28h-9zM39 18h9v28h-9zM25 28h14v8H25z" fill="currentColor" />
                <circle cx="32" cy="14" r="3" fill="currentColor" />
              </svg>
            </div>
            <div className="min-w-0">
              <div className="text-[17px] font-bold leading-tight">
                <span className="text-white">Mali</span><span style={{ color: AMBER }}>Desk</span>
              </div>
              <div className="text-[10px] font-bold tracking-[0.16em] mt-0.5" style={{ color: "#9FC2B4" }}>RENTAL OPERATIONS</div>
            </div>
          </div>
        </button>
      </div>

      <div className="px-5 mt-2">
        <div className="text-[10px] font-bold tracking-wider mb-2" style={{ color: "#7FA396" }}>WORKSPACE</div>
      </div>
      <nav className="px-3 flex flex-col gap-0.5">
        {NAV_ITEMS.filter((item) => !item.adminOnly || role === "Owner/Admin").map((item) => {
          const active = screen === item.key || (screen === "unit" && item.key === "units") || (screen === "statement" && item.key === "units") || (screen === "close" && item.key === "units") || (screen === "archiveDetail" && item.key === "archive");
          return (
            <button
              key={item.key}
              onClick={() => go(item.key)}
              className="flex items-center justify-between px-3 py-2.5 rounded-xl text-sm font-semibold text-left transition-colors"
              style={{ background: active ? SIDEBAR_ACTIVE : "transparent", color: active ? "#fff" : "#B9CCC3" }}
            >
              <span className="flex items-center gap-2.5"><item.icon size={16} />{item.label}</span>
              {active && <ChevronRight size={14} />}
            </button>
          );
        })}
      </nav>

      <div className="mt-auto px-4 pb-4">
        <div className="border-t pt-3" style={{ borderColor: SIDEBAR_BORDER }}>
          <button className="p-1.5 rounded-lg mb-2" style={{ color: "#B9CCC3" }}><Settings size={16} /></button>
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-full flex items-center justify-center text-xs font-bold shrink-0" style={{ background: AMBER, color: SIDEBAR }}>{initials}</div>
            <div className="min-w-0 flex-1">
              <div className="text-sm font-bold truncate" style={{ color: "#fff" }}>{role}</div>
              <div className="text-[11px] truncate" style={{ color: "#8FAB9E" }}>Authenticated account</div>
            </div>
            <button title="Sign out" onClick={onLogout} className="p-1 rounded-lg hover:bg-white/10 shrink-0" style={{ color: "#7FA396" }}><LogOut size={15} /></button>
          </div>
        </div>
      </div>
    </aside>
  );
}

function MobileNav({ screen, go, role }) {
  return (
    <div className="sticky top-0 z-30 flex items-center gap-1 px-2 py-2 overflow-x-auto" style={{ background: SIDEBAR }}>
      {NAV_ITEMS.filter((item) => !item.adminOnly || role === "Owner/Admin").map((item) => {
        const active = screen === item.key;
        return (
          <button key={item.key} onClick={() => go(item.key)} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold whitespace-nowrap"
            style={{ background: active ? SIDEBAR_ACTIVE : "transparent", color: active ? "#fff" : "#B9CCC3" }}>
            <item.icon size={13} />{item.label}
          </button>
        );
      })}
    </div>
  );
}

function TopBar({ q, setQ, onSearchEnter, onBackup }) {
  const [now] = useState(new Date());
  const dateStr = now.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
  return (
    <div className="sticky top-0 z-20 bg-transparent backdrop-blur px-4 sm:px-8 pt-5 pb-3" style={{ background: `${BEIGE}CC` }}>
      <div className="flex items-center justify-between gap-3">
        <div className="relative flex-1 max-w-md">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
          <input
            value={q} onChange={(e) => setQ(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") onSearchEnter(); }}
            placeholder="Search units, tenants, references…"
            className="w-full rounded-xl border pl-9 pr-3 py-2.5 text-sm outline-none bg-white"
            style={{ borderColor: BORDER }}
          />
        </div>
        <div className="flex items-center gap-3">
          <button title="Download backup" onClick={onBackup} className="p-2 rounded-xl bg-white border" style={{ borderColor: BORDER, color: MUTED }}><Download size={16} /></button>
          <button className="p-2 rounded-xl bg-white border" style={{ borderColor: BORDER, color: MUTED }}><Bell size={16} /></button>
          <div className="hidden sm:block text-right leading-tight">
            <div className="text-xs font-semibold" style={{ color: INK }}>{dateStr}</div>
            <div className="text-[11px]" style={{ color: MUTED }}>EAT · Nairobi</div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Main App                                                                 */
/* ---------------------------------------------------------------------- */
function MaliDeskCore({ auth }) {
  const [data, setData] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const role = auth.user.role === "Administrator" ? "Owner/Admin" : auth.user.role === "Manager" ? "Manager" : auth.user.role === "Staff" ? "Staff" : "Read-only viewer";
  const [view, setView] = useState({ screen: "overview" });
  const [toast, setToast] = useState(null);
  const [topQ, setTopQ] = useState("");
  const [quickOpen, setQuickOpen] = useState(false);
  const isDesktop = useIsDesktop(640);

  useEffect(() => {
    (async () => {
      try {
        let loadedData = await storageGet(DATA_KEY);
        if (!loadedData) {
          loadedData = await storageGet(LEGACY_DATA_KEY);
          if (loadedData) {
            loadedData = normaliseData(loadedData);
            await storageSet(DATA_KEY, loadedData);
          }
        }
        setData(normaliseData(loadedData || seedData()));
      } catch {
        setData(normaliseData(seedData()));
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  useEffect(() => {
    if (!loaded || !data) return;
    storageSet(DATA_KEY, normaliseData(data)).catch(() => {});
  }, [data, loaded]);

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2600);
  }, []);

  const rows = useMemo(() => {
    if (!data) return [];
    return data.units.map((u) => {
      const tenant = data.tenants.find((t) => t.id === u.currentTenantId);
      const txns = data.transactions.filter((t) => t.unitId === u.id);
      const balance = u.occupancyStatus === "vacant" ? 0 : unitBalance(txns);
      const status = statusOf(u, balance);
      const priority = priorityOf(status);
      const rank = { HIGH: 4, MEDIUM: 3, LOW: 2, VACANT: 1, PAID: 0 }[priority.key];
      return {
        unit: u, tenant, balance, status, priority, rank,
        depositHeld: u.occupancyStatus === "vacant" ? 0 : u.depositHeld,
        monthsOut: u.occupancyStatus === "vacant" ? 0 : monthsOutstanding(balance, u.rent, u.garbage),
      };
    });
  }, [data]);

  const canEdit = role !== "Read-only viewer";
  const go = (screen, extra = {}) => setView({ screen, ...extra });

  const audit = (action, entity, entityId, details = {}) => ({
    id: uid("audit"),
    at: new Date().toISOString(),
    by: role,
    action,
    entity,
    entityId,
    details,
  });

  if (!loaded || !data) {
    return <div className="min-h-screen flex items-center justify-center text-sm" style={{ color: MUTED, background: BEIGE }}>Loading…</div>;
  }

  const unitTxns = (unitId) => data.transactions.filter((t) => t.unitId === unitId);

  /* ---- mutations ---- */
  function addTransaction(unitId, tx) {
    if (!canEdit) return false;
    const amount = Number(tx.paymentReceived) || 0;
    const rent = Number(tx.rentCharged) || 0;
    const garbage = Number(tx.garbageCharged) || 0;
    if (amount < 0 || rent < 0 || garbage < 0) {
      showToast("Amounts cannot be negative.");
      return false;
    }
    const unitCode = data.units.find((u) => u.id === unitId)?.code || "";
    const transaction = {
      id: uid("tx"),
      unitId,
      createdAt: new Date().toISOString(),
      createdBy: role,
      ...tx,
      paymentReceived: amount,
      rentCharged: rent,
      garbageCharged: garbage,
    };
    transaction.importKey = txImportKey(unitCode, transaction);
    setData((d) => ({
      ...d,
      transactions: [...d.transactions, transaction],
      auditLog: [...d.auditLog, audit("CREATE", "transaction", transaction.id, { unitId, type: tx.type, description: tx.description })],
    }));
    return true;
  }

  /* ---- edit / delete a Payment or Monthly Charge ---- */
  // Full administrator control: any payment or charge — manually entered,
  // imported, or historical — can always be corrected or removed. Editing
  // keeps the same id and marks the record so a later re-import of an old
  // worksheet won't silently overwrite the correction (see importWorkbook).
  function updateTransaction(txId, payload) {
    if (!canEdit) return false;
    const existing = data.transactions.find((t) => t.id === txId);
    if (!existing) { showToast("That transaction could not be found."); return false; }
    if (existing.type !== "payment" && existing.type !== "charge") {
      showToast("Only payments and monthly charges can be edited here.");
      return false;
    }
    const unit = data.units.find((u) => u.id === existing.unitId);

    let updated;
    if (existing.type === "payment") {
      const amount = Number(payload.paymentReceived);
      if (!Number.isFinite(amount) || amount <= 0) {
        showToast("Enter a valid payment amount greater than zero.");
        return false;
      }
      const ref = String(payload.mpesaRef || "").trim();
      if (ref) {
        const duplicate = data.transactions.some(
          (t) => t.id !== txId && String(t.mpesaRef || "").trim().toLowerCase() === ref.toLowerCase()
        );
        if (duplicate) { showToast(`Payment reference ${ref} is already used on another transaction.`); return false; }
      }
      if (!payload.date) { showToast("Date is required."); return false; }
      updated = {
        ...existing,
        date: payload.date, month: monthOf(payload.date),
        paymentReceived: amount, mpesaRef: ref,
        description: String(payload.description || existing.description || "Payment received"),
        remarks: payload.remarks || "",
      };
    } else {
      const rent = Number(payload.rentCharged), garbage = Number(payload.garbageCharged);
      if (!payload.date) { showToast("Date is required."); return false; }
      if (!Number.isFinite(rent) || rent < 0) { showToast("Rent charged cannot be negative."); return false; }
      if (!Number.isFinite(garbage) || garbage < 0) { showToast("Garbage charged cannot be negative."); return false; }
      updated = {
        ...existing,
        date: payload.date, month: monthOf(payload.date),
        rentCharged: rent, garbageCharged: garbage,
        description: String(payload.description || existing.description || "Monthly rent + garbage"),
        remarks: payload.remarks || "",
      };
    }
    updated.updatedAt = new Date().toISOString();
    updated.updatedBy = role;
    updated.editedManually = true; // protects this record from being clobbered by a stale re-import

    setData((d) => ({
      ...d,
      transactions: d.transactions.map((t) => (t.id === txId ? updated : t)),
      auditLog: [
        ...d.auditLog,
        audit("UPDATE", "transaction", txId, {
          unitCode: unit?.code, transactionType: existing.type,
          previous: { date: existing.date, rentCharged: existing.rentCharged, garbageCharged: existing.garbageCharged, paymentReceived: existing.paymentReceived, mpesaRef: existing.mpesaRef, description: existing.description, remarks: existing.remarks },
          updated: { date: updated.date, rentCharged: updated.rentCharged, garbageCharged: updated.garbageCharged, paymentReceived: updated.paymentReceived, mpesaRef: updated.mpesaRef, description: updated.description, remarks: updated.remarks },
        }),
      ],
    }));
    showToast(existing.type === "payment" ? "Payment updated successfully." : "Monthly charge updated successfully.");
    return true;
  }

  function deleteTransaction(txId) {
    if (!canEdit) return false;
    const existing = data.transactions.find((t) => t.id === txId);
    if (!existing) { showToast("That transaction could not be found."); return false; }
    const unit = data.units.find((u) => u.id === existing.unitId);
    const key = existing.importKey || (unit ? txImportKey(unit.code, existing) : null);

    setData((d) => ({
      ...d,
      transactions: d.transactions.filter((t) => t.id !== txId),
      // Tombstone the content fingerprint so a future import of the same old
      // worksheet row doesn't bring this deleted record back to life.
      deletedImportKeys: key && !d.deletedImportKeys.includes(key) ? [...d.deletedImportKeys, key] : d.deletedImportKeys,
      auditLog: [
        ...d.auditLog,
        audit(existing.type === "payment" ? "DELETE_PAYMENT" : "DELETE_CHARGE", "transaction", txId, {
          unitCode: unit?.code, transactionType: existing.type,
          previous: { date: existing.date, rentCharged: existing.rentCharged, garbageCharged: existing.garbageCharged, paymentReceived: existing.paymentReceived, mpesaRef: existing.mpesaRef, description: existing.description, remarks: existing.remarks },
        }),
      ],
    }));
    showToast(existing.type === "payment" ? "Payment deleted." : "Monthly charge deleted.");
    return true;
  }
  function recordPayment(unit, payload) {
    if (!canEdit) return;
    const amount = Number(payload.amount);
    if (!Number.isFinite(amount) || amount <= 0) {
      showToast("Enter a valid payment amount greater than zero.");
      return;
    }
    const ref = String(payload.mpesaRef || "").trim();
    if (ref) {
      const duplicate = data.transactions.some((t) => String(t.mpesaRef || "").trim().toLowerCase() === ref.toLowerCase());
      if (duplicate) {
        showToast(`Payment reference ${ref} has already been recorded.`);
        return;
      }
    }
    const ok = addTransaction(unit.id, {
      date: payload.date || todayISO(), month: monthOf(payload.date), type: "payment",
      description: payload.method ? `Payment (${payload.method})` : "Payment received",
      rentCharged: 0, garbageCharged: 0, paymentReceived: amount,
      mpesaRef: ref, method: payload.method || "Other", remarks: payload.remarks || "",
    });
    if (ok) showToast(`Payment of ${kes(amount)} recorded for ${unit.code}`);
  }
  function addMonthlyCharge(unit, dateForMonth) {
    const month = monthOf(dateForMonth);
    const dup = unitTxns(unit.id).some((t) => t.type === "charge" && t.month === month);
    return { dup, commit: () => {
      if (!canEdit || unit.occupancyStatus === "vacant" || !unit.currentTenantId) {
        showToast("Only occupied units with an active tenant can be charged.");
        return;
      }
      if (dup) {
        showToast(`${unit.code} is already charged for ${monthLabel(month)}.`);
        return;
      }
      const rent = Math.max(0, Number(unit.rent) || 0);
      const garbage = Math.max(0, Number(unit.garbage) || 0);
      addTransaction(unit.id, {
        date: `${month}-01`, month, type: "charge", description: "Monthly rent + garbage",
        rentCharged: rent, garbageCharged: garbage, paymentReceived: 0, mpesaRef: "", remarks: "",
      });
      showToast(`Monthly charge added for ${unit.code} — ${monthLabel(month)}`);
    }};
  }
  function addAdjustment(unit, payload) {
    addTransaction(unit.id, {
      date: payload.date, month: monthOf(payload.date), type: "adjustment", description: payload.description || "Manual adjustment",
      rentCharged: Number(payload.charge) || 0, garbageCharged: 0, paymentReceived: Number(payload.credit) || 0,
      mpesaRef: "", remarks: payload.remarks || "",
    });
    showToast(`Adjustment recorded for ${unit.code}`);
  }
  function closeTenancy(unit, tenant, payload) {
    if (!canEdit) return;
    const refunded = Number(payload.amountRefunded) || 0;
    const forfeited = Number(payload.amountForfeited) || 0;
    if (refunded < 0 || forfeited < 0 || refunded + forfeited > Number(unit.depositHeld || 0)) {
      showToast("Deposit refund plus forfeiture cannot exceed the deposit held.");
      return;
    }
    if (!payload.moveOutDate) {
      showToast("Move-out date is required.");
      return;
    }
    const txns = unitTxns(unit.id);
    const finalBalance = unitBalance(txns);
    setData((d) => ({
      ...d,
      archive: [
        ...d.archive,
        {
          id: uid("arc"), unitId: unit.id, unitCode: unit.code,
          tenant: { ...tenant, moveOutDate: payload.moveOutDate },
          finalBalance, depositHeld: unit.depositHeld,
          amountRefunded: Number(payload.amountRefunded) || 0,
          amountForfeited: Number(payload.amountForfeited) || 0,
          reasonForLeaving: payload.reasonForLeaving, arrearsOutcome: payload.arrearsOutcome,
          evictionNoticeDate: payload.evictionNoticeDate || "", evictionLockDate: payload.evictionLockDate || "",
          recoveryNotes: payload.recoveryNotes || "", notes: payload.notes || "",
          transactions: txns,
        },
      ],
      transactions: d.transactions.filter((t) => t.unitId !== unit.id),
      auditLog: [...d.auditLog, audit("CLOSE", "tenancy", tenant.id, { unitId: unit.id, unitCode: unit.code, finalBalance, refunded, forfeited, moveOutDate: payload.moveOutDate })],
      tenants: d.tenants.filter((t) => t.id !== tenant.id),
      units: d.units.map((u) => (u.id === unit.id ? { ...u, currentTenantId: null, occupancyStatus: "vacant", depositHeld: 0 } : u)),
    }));
    showToast(`Tenancy closed for ${unit.code}. Unit is now vacant.`);
    go("units");
  }

  function updateUnitCharges(unitId, rent, garbage) {
    if (!canEdit) return false;
    const r = Number(rent), g = Number(garbage);
    if (!Number.isFinite(r) || r < 0 || !Number.isFinite(g) || g < 0) {
      showToast("Rent and garbage must be valid non-negative amounts.");
      return false;
    }
    const unit = data.units.find((u) => u.id === unitId);
    if (!unit) return false;
    setData((d) => ({
      ...d,
      units: d.units.map((u) => (u.id === unitId ? { ...u, rent: r, garbage: g } : u)),
      auditLog: [...d.auditLog, audit("UPDATE", "unit_charges", unitId, { unitCode: unit.code, from: { rent: unit.rent, garbage: unit.garbage }, to: { rent: r, garbage: g } })],
    }));
    showToast(`${unit.code} charges updated successfully.`);
    return true;
  }

  function addUnit(payload) {
    if (!canEdit) return { error: "You do not have permission to add units." };
    const code = String(payload.code || "").trim();
    if (!code) return { error: "Unit Code is required." };
    const exists = data.units.some((u) => u.code.trim().toLowerCase() === code.toLowerCase());
    if (exists) return { error: `A unit with code ${code} already exists.` };
    const rent = Number(payload.rent), garbage = Number(payload.garbage);
    if (!Number.isFinite(rent) || rent < 0) return { error: "Rent cannot be negative." };
    if (!Number.isFinite(garbage) || garbage < 0) return { error: "Garbage cannot be negative." };

    const unitId = uid("u");
    const now = new Date().toISOString();

    if (payload.addTenantNow) {
      const tenantId = uid("t");
      const deposit = Number(payload.deposit);
      if (!Number.isFinite(deposit) || deposit < 0) return { error: "Deposit cannot be negative." };
      if (!String(payload.tenantName || "").trim()) return { error: "Tenant name is required." };
      if (!payload.moveInDate) return { error: "Move-in date is required." };

      const depositTx = deposit > 0 ? [{
        id: uid("tx"), unitId, date: payload.moveInDate, month: monthOf(payload.moveInDate), type: "deposit",
        description: "Security Deposit Received", rentCharged: 0, garbageCharged: 0, paymentReceived: 0,
        mpesaRef: "", remarks: `Deposit of ${kes(deposit)} received`, createdBy: role, createdAt: now,
      }] : [];
      const chargeTx = payload.createFirstCharge ? [{
        id: uid("tx"), unitId, date: payload.moveInDate, month: monthOf(payload.moveInDate), type: "charge",
        description: "Monthly rent + garbage", rentCharged: rent, garbageCharged: garbage, paymentReceived: 0,
        mpesaRef: "", remarks: "", createdBy: role, createdAt: now,
      }] : [];

      setData((d) => ({
        ...d,
        units: [...d.units, { id: unitId, code, type: payload.type, rent, garbage, depositHeld: deposit, currentTenantId: tenantId, occupancyStatus: "occupied" }],
        tenants: [...d.tenants, { id: tenantId, name: payload.tenantName.trim(), phone: payload.phone || "", email: payload.email || "", moveInDate: payload.moveInDate, moveOutDate: null, notes: payload.notes || "" }],
        transactions: [...d.transactions, ...depositTx, ...chargeTx],
        auditLog: [...d.auditLog, audit("CREATE", "unit", unitId, { unitCode: code, tenantId, rent, garbage, deposit })],
      }));
      showToast(`Unit ${code} created and tenant added successfully.`);
    } else {
      setData((d) => ({
        ...d,
        units: [...d.units, { id: unitId, code, type: payload.type, rent, garbage, depositHeld: 0, currentTenantId: null, occupancyStatus: "vacant" }],
        auditLog: [...d.auditLog, audit("CREATE", "unit", unitId, { unitCode: code, rent, garbage, vacant: true })],
      }));
      showToast(`Unit ${code} created successfully.`);
    }
    return { unitId };
  }

  function addTenantToUnit(unitId, payload) {
    if (!canEdit) return { error: "You do not have permission to add tenants." };
    const unit = data.units.find((u) => u.id === unitId);
    if (!unit) return { error: "Unit not found." };
    if (unit.occupancyStatus !== "vacant" || unit.currentTenantId) return { error: "This unit is already occupied." };

    const name = String(payload.name || "").trim();
    const rent = Number(payload.rent), garbage = Number(payload.garbage), deposit = Number(payload.deposit ?? 0);
    if (!name) return { error: "Tenant name is required." };
    if (!payload.moveInDate) return { error: "Move-in date is required." };
    if (!Number.isFinite(rent) || rent < 0) return { error: "Monthly rent must be valid and non-negative." };
    if (!Number.isFinite(garbage) || garbage < 0) return { error: "Monthly garbage must be valid and non-negative." };
    if (!Number.isFinite(deposit) || deposit < 0) return { error: "Deposit cannot be negative." };

    const tenantId = uid("t");
    const now = new Date().toISOString();
    const depositTx = deposit > 0 ? [{
      id: uid("tx"), unitId, date: payload.moveInDate, month: monthOf(payload.moveInDate), type: "deposit",
      description: "Security Deposit Received", rentCharged: 0, garbageCharged: 0, paymentReceived: 0,
      mpesaRef: "", remarks: `Deposit of ${kes(deposit)} received`, createdBy: role, createdAt: now,
    }] : [];
    const chargeTx = payload.createFirstCharge ? [{
      id: uid("tx"), unitId, date: payload.moveInDate, month: monthOf(payload.moveInDate), type: "charge",
      description: "Monthly rent + garbage", rentCharged: rent, garbageCharged: garbage, paymentReceived: 0,
      mpesaRef: "", remarks: "", createdBy: role, createdAt: now,
    }] : [];

    setData((d) => ({
      ...d,
      tenants: [...d.tenants, { id: tenantId, name, phone: payload.phone || "", email: payload.email || "", moveInDate: payload.moveInDate, moveOutDate: null, notes: payload.notes || "" }],
      units: d.units.map((u) => (u.id === unitId ? { ...u, currentTenantId: tenantId, occupancyStatus: "occupied", rent, garbage, depositHeld: deposit } : u)),
      transactions: [...d.transactions, ...depositTx, ...chargeTx],
      auditLog: [...d.auditLog, audit("CREATE", "tenant", tenantId, { unitId, name, rent, garbage, deposit })],
    }));
    showToast(`Tenant added to ${unit.code} successfully.`);
    return { tenantId };
  }

  function updateTenant(tenantId, payload) {
    if (!canEdit) return false;
    const name = String(payload.name || "").trim();
    if (!name || !payload.moveInDate) {
      showToast("Tenant name and move-in date are required.");
      return false;
    }
    setData((d) => ({
      ...d,
      tenants: d.tenants.map((t) => (t.id === tenantId ? { ...t, name, phone: payload.phone || "", email: payload.email || "", moveInDate: payload.moveInDate, notes: payload.notes || "" } : t)),
      auditLog: [...d.auditLog, audit("UPDATE", "tenant", tenantId, { name, phone: payload.phone || "", email: payload.email || "" })],
    }));
    showToast(`Tenant details updated successfully.`);
    return true;
  }


  /* ---- import (template-matched Excel) ---- */
  // Full-replace import: whatever is parsed from the workbook REPLACES the
  // current units, tenants and transactions entirely — nothing from the
  // sheets is merged into what's already stored. Archive/audit log/settings
  // are untouched, since the template doesn't carry those. The caller
  // (ImportExport) is expected to confirm with the user before calling this,
  // since it's destructive and can't be undone from within the app.
  async function importWorkbook(file) {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: false });
      // Skip the source template's non-ledger tabs, AND the sheet names
      // this app's own Export produces — so importing a file you just
      // exported from MaliDesk never gets misread as per-unit ledgers.
      const skipSheets = new Set(["Summary", "Statement", "Archive", "Instructions", "Units & Balances", "Transactions", "Closed Tenancies", "Empty"]);
      const sheetNames = wb.SheetNames.filter((n) => !skipSheets.has(n));
      const get = (ws, r, c) => { const cell = ws[XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })]; return cell ? cell.v : undefined; };
      const toISO = (v) => {
          if (v == null || v === "") return null;
          const pad = (n) => String(n).padStart(2, "0");

          // Excel serial number (most reliable – no timezone issues)
          if (typeof v === "number" && Number.isFinite(v)) {
            const parsed = XLSX.SSF.parse_date_code(v);
            if (parsed?.y && parsed?.m && parsed?.d) {
              return `${parsed.y}-${pad(parsed.m)}-${pad(parsed.d)}`;
            }
          }

          // Fallback for Date objects (use UTC to avoid off-by-one)
          if (v instanceof Date && !Number.isNaN(v.getTime())) {
            return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
          }

          // Text values
          const text = String(v).trim();
          if (!text) return null;

          const iso = text.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
          if (iso) return `${iso[1]}-${pad(iso[2])}-${pad(iso[3])}`;

          const dmy = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
          if (dmy) return `${dmy[3]}-${pad(dmy[2])}-${pad(dmy[1])}`;

          return text;
        };
      // A cell only counts as filled once null/undefined, empty strings and
      // whitespace-only strings are all ruled out — this is what keeps
      // unused ledger rows and blank formula results out of the import.
      const hasValue = (v) => v !== null && v !== undefined && String(v).trim() !== "";

      // Prior state, indexed by content fingerprint (unit CODE, not the
      // soon-to-be-replaced internal unitId) so edits and deletions can be
      // recognised across this full-replace import:
      //  - a fingerprint in deletedKeys means the user deleted that exact
      //    row before — don't let this import bring it back.
      //  - a fingerprint matching a transaction flagged editedManually means
      //    the user corrected that row before — keep the correction instead
      //    of overwriting it with the sheet's (now-stale) values.
      const unitCodeById = new Map(data.units.map((u) => [u.id, u.code]));
      const priorByKey = new Map();
      data.transactions.forEach((t) => {
        const code = unitCodeById.get(t.unitId);
        if (!code) return;
        priorByKey.set(t.importKey || txImportKey(code, t), t);
      });
      const deletedKeys = new Set(data.deletedImportKeys || []);

      const newUnits = [], newTenants = [], newTxns = [];
      let unitCount = 0, txnCount = 0, skippedRowCount = 0, skippedSheetCount = 0, resurrectionsBlocked = 0, editsPreserved = 0;

      for (const name of sheetNames) {
        const ws = wb.Sheets[name];
        const code = get(ws, 3, 2);
        if (!code) continue;
        // Extra safety net beyond the skip-list: a real unit ledger has
        // "Date" as its row-5 column-1 header. Any other sheet shape (a
        // report, a summary, anything not built from this template) gets
        // skipped here instead of being misread as ledger rows.
        const headerCell = get(ws, 5, 1);
        if (!headerCell || !/^date$/i.test(String(headerCell).trim())) { skippedSheetCount++; continue; }

        // Pass 1: collect valid rows (Date AND Description both filled).
        // Anything else — fully blank rows, or a row with only one of the
        // two — is overlooked, not imported, and never treated as new data.
        const rows = [];
        let r = 6;
        let blankStreak = 0;
        while (r < 400) {
          const dateCell = get(ws, r, 1), descCell = get(ws, r, 3);
          const rowHasDate = hasValue(dateCell);
          const rowHasDesc = hasValue(descCell);
          if (!rowHasDate && !rowHasDesc) {
            blankStreak++;
            if (blankStreak >= 2) break; // two blank rows in a row = end of this unit's table
            skippedRowCount++;
            r++;
            continue;
          }
          blankStreak = 0;
          if (rowHasDate && rowHasDesc) {
            rows.push({
              date: toISO(dateCell), description: String(descCell).trim(),
              rentCharged: get(ws, r, 5), garbageCharged: get(ws, r, 6), paymentReceived: get(ws, r, 7),
              mpesaRef: get(ws, r, 8), remarks: get(ws, r, 10),
            });
          } else {
            skippedRowCount++; // date-only or description-only row
          }
          r++;
        }

        // Every matched sheet becomes a fresh unit + tenant, built purely
        // from what's on the sheet — no lookup against existing data, since
        // this is a full replace, not a merge.
        const rent = get(ws, 3, 4) || 0;
        const garbage = get(ws, 3, 6) || 0;
        const tenantName = get(ws, 3, 10) || "Unnamed tenant";
        const deposit = get(ws, 3, 12) || 0;
        const unitId = uid("u");
        const tenantId = uid("t");
        const type = String(code).toLowerCase().startsWith("shop") ? "shop" : "residential";
        const dates = rows.map((t) => t.date).filter(Boolean).sort();
        newUnits.push({ id: unitId, code: String(code), type, rent, garbage, depositHeld: deposit, currentTenantId: tenantId, occupancyStatus: "occupied" });
        newTenants.push({ id: tenantId, name: String(tenantName), phone: "", email: "", moveInDate: dates[0] || null, moveOutDate: null, notes: "Imported from Excel." });
        unitCount++;

        for (const row of rows) {
          const dateVal = row.date;
          const type2 = /charge/i.test(row.description) ? "charge" : "payment";
          const rowContent = {
            type: type2,
            date: dateVal,
            description: hasValue(row.description) ? String(row.description).trim() : "",
            rentCharged: hasValue(row.rentCharged) ? Number(row.rentCharged) : 0,
            garbageCharged: hasValue(row.garbageCharged) ? Number(row.garbageCharged) : 0,
            paymentReceived: hasValue(row.paymentReceived) ? Number(row.paymentReceived) : 0,
            mpesaRef: hasValue(row.mpesaRef) ? String(row.mpesaRef).trim() : "",
          };
          const key = txImportKey(code, rowContent);

          if (deletedKeys.has(key)) {
            // This exact row was deliberately deleted before — the sheet
            // still has it, but it must not come back.
            resurrectionsBlocked++;
            continue;
          }

          const prior = priorByKey.get(key);
          if (prior && prior.editedManually) {
            // Keep the user's correction, just re-pointed at this import's
            // (freshly generated) unit id so it still shows on the right
            // ledger.
            newTxns.push({ ...prior, unitId });
            editsPreserved++;
            txnCount++;
            continue;
          }

          newTxns.push({
            id: uid("tx"), unitId, date: dateVal, month: dateVal ? dateVal.slice(0, 7) : null, type: type2,
            description: rowContent.description,
            rentCharged: rowContent.rentCharged, garbageCharged: rowContent.garbageCharged, paymentReceived: rowContent.paymentReceived,
            mpesaRef: rowContent.mpesaRef,
            remarks: hasValue(row.remarks) ? String(row.remarks).trim() : "",
            createdBy: "Import", createdAt: dateVal,
            importKey: key,
          });
          txnCount++;
        }
      }

      if (unitCount === 0) {
        showToast("Could not read that file — no unit sheets matched the expected layout. Nothing was changed.");
        return;
      }

      // Full replace: the file's contents become the new units/tenants/
      // transactions in their entirety. Archive, audit log, settings and the
      // deleted-record tombstones are left as they were, since nothing in
      // the template describes them.
      setData((d) => ({ ...d, units: newUnits, tenants: newTenants, transactions: newTxns }));

      let msg = `Import complete: replaced existing data with ${unitCount} unit(s) and ${txnCount} transaction(s) from the file.`;
      if (editsPreserved) msg += ` ${editsPreserved} previously-edited record(s) kept as you corrected them.`;
      if (resurrectionsBlocked) msg += ` ${resurrectionsBlocked} row(s) skipped — they match a transaction you deleted.`;
      if (skippedRowCount) msg += ` ${skippedRowCount} incomplete row(s) overlooked (missing date or description).`;
      if (skippedSheetCount) msg += ` ${skippedSheetCount} sheet(s) skipped (unexpected layout).`;
      showToast(msg);
    } catch (e) {
      showToast("Could not read that file — check it matches the expected sheet layout.");
    }
  }

  /* ---- export ---- */
  function buildWorkbook(include) {
    const wb = XLSX.utils.book_new();
    if (include.units) {
      const rows = data.units.map((u) => {
        const tenant = data.tenants.find((t) => t.id === u.currentTenantId);
        const bal = unitBalance(unitTxns(u.id));
        const status = statusOf(u, bal);
        const pr = priorityOf(status);
        return {
          Unit: u.code, Type: u.type, Tenant: tenant ? tenant.name : "Vacant",
          "Rent (KES)": u.rent, "Garbage (KES)": u.garbage,
          "Deposit Held (KES)": u.occupancyStatus === "vacant" ? 0 : u.depositHeld,
          "Current Balance (KES)": u.occupancyStatus === "vacant" ? 0 : bal,
          "Months Outstanding": u.occupancyStatus === "vacant" ? 0 : Math.round(monthsOutstanding(bal, u.rent, u.garbage) * 100) / 100,
          Status: status, Priority: pr.key, "Suggested Action": pr.action,
        };
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Units & Balances");
    }
    if (include.transactions) {
      const txRows = data.transactions.slice().sort((a, b) => String(a.date || "").localeCompare(String(b.date || ""))).map((t) => {
        const u = data.units.find((u) => u.id === t.unitId);
        return {
          Unit: u?.code || "", Date: t.date, Month: monthLabel(t.month), Type: t.type, Description: t.description,
          "Rent Charged": t.rentCharged, "Garbage Charged": t.garbageCharged, "Payment Received": t.paymentReceived,
          "Payment Method": t.method || "", "M-Pesa Ref": t.mpesaRef, Remarks: t.remarks,
        };
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(txRows.length ? txRows : [{ Note: "No transactions yet" }]), "Transactions");
    }
    if (include.archive) {
      const archiveRows = data.archive.map((a) => ({
        Unit: a.unitCode, Tenant: a.tenant.name, Phone: a.tenant.phone,
        "Move-in": a.tenant.moveInDate, "Move-out": a.tenant.moveOutDate, "Final Balance (KES)": a.finalBalance,
        "Deposit Held (KES)": a.depositHeld, "Amount Refunded (KES)": a.amountRefunded, "Amount Forfeited (KES)": a.amountForfeited,
        "Reason for Leaving": a.reasonForLeaving, "Arrears Outcome": a.arrearsOutcome, Notes: a.notes,
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(archiveRows.length ? archiveRows : [{ Note: "No closed tenancies yet" }]), "Closed Tenancies");
    }
    if (wb.SheetNames.length === 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ Note: "Nothing selected" }]), "Empty");
    return wb;
  }
  function exportExcel(include) {
    const wb = buildWorkbook(include);
    XLSX.writeFile(wb, `malidesk-export-${todayISO()}.xlsx`);
    showToast("Excel export downloaded");
  }
  function exportCSV(include) {
    const wb = buildWorkbook(include);
    const firstSheet = wb.SheetNames[0];
    const csv = XLSX.utils.sheet_to_csv(wb.Sheets[firstSheet]);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `malidesk-${firstSheet.toLowerCase().replace(/\s+/g, "-")}-${todayISO()}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    showToast(`CSV downloaded (${firstSheet})`);
  }

  return (
    <div className={isDesktop ? "min-h-screen flex" : "min-h-screen"} style={{ background: BEIGE, color: INK }}>
      <style>{`.tabnum{font-variant-numeric:tabular-nums}`}</style>
      {isDesktop ? (
        <Sidebar screen={view.screen} go={go} role={role} onLogout={auth.logout} />
      ) : (
        <MobileNav screen={view.screen} go={go} role={role} />
      )}
      <div className="flex-1 min-w-0">
        <div className="bg-amber-50 border-b px-4 sm:px-8 py-1.5 text-[11px] text-amber-800 flex items-center justify-between gap-2">
          <span><strong>Local data protection:</strong> This build includes resilient browser storage and manual JSON backups.</span>
          {isDesktop && <span>For multi-device production use, connect the data layer to a secure cloud database.</span>}
        </div>
        <TopBar
          q={topQ}
          setQ={setTopQ}
          onSearchEnter={() => go("units")}
          onBackup={() => {
            const blob = new Blob([JSON.stringify({ ...normaliseData(data), exportedAt: new Date().toISOString(), app: "MaliDesk" }, null, 2)], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `malidesk-backup-${todayISO()}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast("Backup downloaded");
          }}
        />
        <main className="px-4 sm:px-8 pb-16 pt-1">
          {view.screen === "overview" && (
            <Overview
              rows={rows} data={data} canEdit={canEdit} role={role}
              openUnit={(id) => go("unit", { unitId: id })} goUnits={() => go("units")}
              onExport={() => exportExcel({ units: true, transactions: true, archive: true })}
              onQuickAction={() => setQuickOpen(true)}
            />
          )}
          {view.screen === "units" && (
            <UnitsScreen rows={rows} data={data} canEdit={canEdit} openUnit={(id) => go("unit", { unitId: id })} initialQuery={topQ} goImportExport={() => go("importexport")} onAddUnit={addUnit} />
          )}
          {view.screen === "unit" && (
            <UnitDetail
              data={data} unitId={view.unitId} canEdit={canEdit}
              back={() => go("units")}
              onStatement={() => go("statement", { unitId: view.unitId })}
              onCloseTenancy={() => go("close", { unitId: view.unitId })}
              recordPayment={recordPayment} addMonthlyCharge={addMonthlyCharge} addAdjustment={addAdjustment}
              updateUnitCharges={updateUnitCharges} addTenantToUnit={addTenantToUnit} updateTenant={updateTenant}
              updateTransaction={updateTransaction} deleteTransaction={deleteTransaction}
            />
          )}
          {view.screen === "statement" && (
            <Statement data={data} unitId={view.unitId} back={() => go("unit", { unitId: view.unitId })} />
          )}
          {view.screen === "close" && (
            <CloseTenancyFlow data={data} unitId={view.unitId} back={() => go("unit", { unitId: view.unitId })} onSubmit={closeTenancy} />
          )}
          {view.screen === "archive" && (
            <ArchiveList data={data} openDetail={(id) => go("archiveDetail", { archiveId: id })} />
          )}
          {view.screen === "archiveDetail" && (
            <ArchiveDetail data={data} archiveId={view.archiveId} back={() => go("archive")} />
          )}
          {view.screen === "users" && role === "Owner/Admin" && (
            <UserManagement auth={auth} />
          )}
          {view.screen === "importexport" && (
            <ImportExport
              data={data} onImport={importWorkbook} onExportExcel={exportExcel} onExportCSV={exportCSV}
              showColumns={() => showToast("Row 3: unit code, rent, garbage, tenant name, deposit. Row 6+: ledger — date, description, rent charged, garbage charged, payment, M-Pesa ref, remarks.")}
            />
          )}
        </main>
      </div>
      {quickOpen && (
        <QuickActionModal data={data} onClose={() => setQuickOpen(false)} onGo={(unitId, action) => { setQuickOpen(false); go("unit", { unitId }); }} />
      )}
      {toast && (
        <div className="fixed bottom-4 left-1/2 -translate-x-1/2 bg-black text-white text-sm px-4 py-2.5 rounded-full shadow-lg z-50">
          {toast}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Overview                                                                 */
/* ---------------------------------------------------------------------- */
function Overview({ rows, data, canEdit, openUnit, goUnits, onExport, onQuickAction }) {
  const occupied = rows.filter((r) => r.unit.occupancyStatus !== "vacant");
  const vacant = rows.filter((r) => r.unit.occupancyStatus === "vacant");

  const attention = rows.filter((r) => r.priority.key === "HIGH" || r.priority.key === "MEDIUM" || r.priority.key === "LOW").sort((a, b) => b.rank - a.rank || b.balance - a.balance).slice(0, 6);

  const recent = [...data.transactions].sort((a, b) => (String(b.date || "") + String(b.createdAt || "")).localeCompare(String(a.date || "") + String(a.createdAt || ""))).slice(0, 6);

  const mixLabels = { HIGH: "In arrears", MEDIUM: "Follow up", LOW: "Monitor", PAID: "Up to date", VACANT: "Vacant" };
  const mix = ["PAID", "LOW", "MEDIUM", "HIGH", "VACANT"].map((k) => ({ key: k, count: rows.filter((r) => r.priority.key === k).length }));

  const now = new Date();
  const dateLabel = now.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).toUpperCase().replace(/,/, " ·");

  return (
    <div>
      <div className="flex flex-wrap items-start justify-between gap-3 mb-6 pt-2">
        <div>
          <div className="text-xs font-bold tracking-wide mb-1.5" style={{ color: MUTED }}>{dateLabel}</div>
          <h1 className="text-xl sm:text-2xl font-bold">{greeting()}.</h1>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>Here's the pulse of your rental portfolio.</p>
        </div>
        <div className="flex gap-2">
          <Btn variant="secondary" onClick={onExport}>Monthly report</Btn>
          {canEdit && <Btn onClick={onQuickAction} icon={Plus}>Record activity</Btn>}
        </div>
      </div>


      <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
        <div className="lg:col-span-2">
          <div className="rounded-2xl border bg-white overflow-hidden mb-6" style={{ borderColor: BORDER }}>
            <div className="flex items-center justify-between px-5 pt-4 pb-3">
              <div>
                <div className="text-[11px] font-bold tracking-wide" style={{ color: MUTED }}>ATTENTION QUEUE</div>
                <div className="text-sm font-bold mt-0.5">Balances worth a look</div>
              </div>
              <button onClick={goUnits} className="text-xs font-semibold flex items-center gap-1 shrink-0" style={{ color: "#1F8A4C" }}>View ledger <ChevronRight size={13} /></button>
            </div>
            <div>
              {attention.map((r) => {
                const c = STATUS_COLORS[r.priority.key];
                const label = r.priority.key === "HIGH" ? "In Arrears" : r.priority.key === "MEDIUM" ? "Review" : "Monitor";
                return (
                  <button key={r.unit.id} onClick={() => openUnit(r.unit.id)} className="w-full text-left flex items-center gap-3 px-5 py-3 border-t hover:bg-black/[0.015]" style={{ borderColor: BORDER }}>
                    <div className="w-11 h-11 rounded-xl flex items-center justify-center text-xs font-bold shrink-0" style={{ background: c.bg, color: c.fg }}>{r.unit.code}</div>
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-2">
                        <span className="font-semibold text-sm truncate">{r.tenant?.name || "—"}</span>
                        <Pill label={label} colorKey={r.priority.key} />
                      </div>
                      <div className="text-xs truncate mt-0.5" style={{ color: MUTED }}>{r.unit.code} · {r.priority.action}</div>
                    </div>
                    <div className="text-right shrink-0">
                      <div className="text-sm font-bold tabnum" style={{ color: "#C0392B" }}>{kes(r.balance)}</div>
                      <div className="text-[11px] mt-0.5" style={{ color: MUTED }}>{r.monthsOut.toFixed(2)} months</div>
                    </div>
                    <ChevronRight size={16} style={{ color: MUTED }} className="shrink-0" />
                  </button>
                );
              })}
              {attention.length === 0 && (
                <div className="px-5 py-8 text-sm text-center border-t" style={{ borderColor: BORDER, color: MUTED }}>
                  Nothing needs attention right now — every tenant is paid up or vacant.
                </div>
              )}
            </div>
          </div>

          <div className="flex items-center justify-between mb-3">
            <div className="text-xs font-bold tracking-wide" style={{ color: MUTED }}>RECENT ACTIVITY</div>
            <button onClick={goUnits} className="text-xs font-semibold flex items-center gap-1" style={{ color: "#1F8A4C" }}>Full ledger <ChevronRight size={13} /></button>
          </div>
          <div className="rounded-2xl border bg-white overflow-hidden" style={{ borderColor: BORDER }}>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead><tr style={{ borderBottom: `1px solid ${BORDER}` }}>
                  {["Date", "Description", "Unit", "Received", "Balance"].map((h) => <th key={h} className="text-left px-4 py-2.5 text-xs font-bold" style={{ color: MUTED }}>{h}</th>)}
                </tr></thead>
                <tbody>
                  {recent.map((t) => {
                    const u = data.units.find((u) => u.id === t.unitId);
                    return (
                      <tr key={t.id} style={{ borderBottom: `1px solid ${BORDER}` }}>
                        <td className="px-4 py-2.5 whitespace-nowrap">{fmtDate(t.date)}</td>
                        <td className="px-4 py-2.5 whitespace-nowrap">{t.description}</td>
                        <td className="px-4 py-2.5 whitespace-nowrap font-semibold">{u?.code}</td>
                        <td className="px-4 py-2.5 tabnum whitespace-nowrap" style={{ color: t.paymentReceived ? "#1F8A4C" : MUTED }}>{t.paymentReceived ? kes(t.paymentReceived) : "—"}</td>
                        <td className="px-4 py-2.5 tabnum whitespace-nowrap" style={{ color: t.rentCharged || t.garbageCharged ? "#C0392B" : MUTED }}>{(t.rentCharged || t.garbageCharged) ? kes(t.rentCharged + t.garbageCharged) : "—"}</td>
                      </tr>
                    );
                  })}
                  {recent.length === 0 && <tr><td colSpan={5} className="px-4 py-8 text-center" style={{ color: MUTED }}>No activity yet.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <div>
          <div className="rounded-2xl border bg-white overflow-hidden" style={{ borderColor: BORDER }}>
            <div className="px-5 pt-4 pb-3">
              <div className="text-[11px] font-bold tracking-wide" style={{ color: MUTED }}>PORTFOLIO MIX</div>
              <div className="text-sm font-bold mt-0.5">At a glance</div>
            </div>
            {mix.map((m) => (
              <div key={m.key} className="flex items-center justify-between px-5 py-2.5 border-t" style={{ borderColor: BORDER }}>
                <span className="flex items-center gap-2 text-sm font-medium">
                  <span className="w-2 h-2 rounded-full shrink-0" style={{ background: STATUS_COLORS[m.key].dot }} />
                  {mixLabels[m.key]}
                </span>
                <span className="text-sm tabnum" style={{ color: MUTED }}>{m.count} unit{m.count === 1 ? "" : "s"}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function QuickActionModal({ data, onClose, onGo }) {
  const [unitId, setUnitId] = useState(data.units[0]?.id || "");
  return (
    <Modal title="Record activity" onClose={onClose}>
      <Field label="Unit">
        <select value={unitId} onChange={(e) => setUnitId(e.target.value)} className={inputCls} style={inputStyle}>
          {data.units.map((u) => <option key={u.id} value={u.id}>{u.code}{u.occupancyStatus === "vacant" ? " (vacant)" : ""}</option>)}
        </select>
      </Field>
      <p className="text-xs mb-4" style={{ color: MUTED }}>You'll land on the unit page to record a payment, add a charge, or make an adjustment.</p>
      <div className="flex justify-end gap-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => onGo(unitId)}>Open Unit</Btn>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/* Units & Balances                                                         */
/* ---------------------------------------------------------------------- */
function UnitsScreen({ rows, data, canEdit, openUnit, initialQuery, goImportExport, onAddUnit }) {
  const [q, setQ] = useState(initialQuery || "");
  const [typeFilter, setTypeFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [priorityFilter, setPriorityFilter] = useState("all");
  const [sort, setSort] = useState({ key: "unit", dir: "asc" });
  const [addUnitOpen, setAddUnitOpen] = useState(false);

  const filtered = rows.filter((r) => {
    if (typeFilter !== "all" && r.unit.type !== typeFilter) return false;
    if (statusFilter !== "all" && r.status !== statusFilter) return false;
    if (priorityFilter !== "all" && r.priority.key !== priorityFilter) return false;
    if (q) {
      const hay = `${r.unit.code} ${r.tenant?.name || ""} ${r.tenant?.phone || ""} ${r.tenant?.email || ""}`.toLowerCase();
      if (!hay.includes(q.toLowerCase())) return false;
    }
    return true;
  });
  const sorted = [...filtered].sort((a, b) => {
    let av, bv;
    switch (sort.key) {
      case "unit": av = a.unit.code; bv = b.unit.code; break;
      case "tenant": av = a.tenant?.name || ""; bv = b.tenant?.name || ""; break;
      case "balance": av = a.balance; bv = b.balance; break;
      case "rent": av = a.unit.rent; bv = b.unit.rent; break;
      default: av = a.rank; bv = b.rank;
    }
    if (av < bv) return sort.dir === "asc" ? -1 : 1;
    if (av > bv) return sort.dir === "asc" ? 1 : -1;
    return 0;
  });
  const toggleSort = (key) => setSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "asc" }));
  const reset = () => { setQ(""); setTypeFilter("all"); setStatusFilter("all"); setPriorityFilter("all"); setSort({ key: "unit", dir: "asc" }); };

  const Th = ({ label, k }) => (
    <th onClick={() => toggleSort(k)} className="text-left px-3 py-2.5 text-xs font-bold cursor-pointer select-none whitespace-nowrap" style={{ color: MUTED }}>
      <span className="inline-flex items-center gap-1">{label}{sort.key === k && (sort.dir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}</span>
    </th>
  );

  return (
    <div>
      <div className="flex flex-wrap items-start justify-between gap-3 mb-1 pt-2">
        <div>
          <div className="text-[11px] font-bold tracking-wide mb-1" style={{ color: MUTED }}>CURRENT PORTFOLIO</div>
          <h1 className="text-xl font-bold">Units &amp; balances</h1>
          <p className="text-sm mt-0.5" style={{ color: MUTED }}>A clean ledger of who owes what, right now.</p>
        </div>
        <div className="flex items-center gap-2">
          <span className="text-xs font-semibold rounded-full px-3 py-1.5" style={{ background: "#EFEAE0", color: MUTED }}>{filtered.length} of {rows.length} units</span>
          <Btn variant="secondary" size="sm" icon={Filter} onClick={reset}>Reset</Btn>
          {canEdit && <Btn size="sm" onClick={() => setAddUnitOpen(true)}>+ Add Unit</Btn>}
        </div>
      </div>

      <div className="rounded-2xl border p-3 my-4 flex flex-wrap gap-2 items-center bg-white" style={{ borderColor: BORDER }}>
        <div className="relative flex-1 min-w-[200px]">
          <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search by unit, tenant or phone…" className={`${inputCls} pl-8`} style={inputStyle} />
        </div>
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="text-xs rounded-xl border px-3 py-2 font-semibold" style={inputStyle}>
          <option value="all">All statuses</option>
          {["Vacant", "Paid/Credit", "<1 month", "1–2 months", "2–3 months", ">3 months"].map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <select value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)} className="text-xs rounded-xl border px-2.5 py-2" style={inputStyle}>
          <option value="all">All types</option><option value="residential">Residential</option><option value="shop">Shop</option>
        </select>
        <select value={priorityFilter} onChange={(e) => setPriorityFilter(e.target.value)} className="text-xs rounded-xl border px-2.5 py-2" style={inputStyle}>
          <option value="all">All priorities</option>
          {[["VACANT", "Vacant"], ["PAID", "Paid"], ["LOW", "Low"], ["MEDIUM", "Medium"], ["HIGH", "High"]].map(([v, l]) => <option key={v} value={v}>{l}</option>)}
        </select>
      </div>

      <div className="rounded-2xl border overflow-hidden bg-white" style={{ borderColor: BORDER }}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead style={{ borderBottom: `1px solid ${BORDER}` }}>
              <tr>
                {Th({ label: "Unit", k: "unit" })}{Th({ label: "Tenant", k: "tenant" })}
                <th className="text-left px-3 py-2.5 text-xs font-bold" style={{ color: MUTED }}>Monthly charge</th>
                {Th({ label: "Balance", k: "balance" })}
                <th className="text-left px-3 py-2.5 text-xs font-bold" style={{ color: MUTED }}>Status</th>
                <th className="text-left px-3 py-2.5 text-xs font-bold" style={{ color: MUTED }}>Suggested next step</th>
                <th className="text-left px-3 py-2.5 text-xs font-bold" style={{ color: MUTED }}></th>
              </tr>
            </thead>
            <tbody>
              {sorted.map((r) => {
                const c = STATUS_COLORS[r.priority.key];
                return (
                  <tr key={r.unit.id} onClick={() => openUnit(r.unit.id)} className="cursor-pointer hover:bg-black/[0.015]" style={{ borderBottom: `1px solid ${BORDER}` }}>
                    <td className="px-3 py-3 whitespace-nowrap">
                      <div className="flex items-center gap-2.5">
                        <div className="w-9 h-9 rounded-lg flex items-center justify-center text-[11px] font-bold shrink-0" style={{ background: c.bg, color: c.fg }}>{r.unit.code}</div>
                        <div className="text-[11px]" style={{ color: MUTED }}>{r.unit.type === "shop" ? "Shop" : "Residential"}</div>
                      </div>
                    </td>
                    <td className="px-3 py-3 whitespace-nowrap">
                      {r.tenant ? <div className="font-medium">{r.tenant.name}</div> : <span style={{ color: MUTED }}>Vacant</span>}
                      {r.tenant?.phone && <div className="text-[11px]" style={{ color: MUTED }}>{r.tenant.phone}</div>}
                    </td>
                    <td className="px-3 py-3 whitespace-nowrap">
                      <div className="tabnum font-medium">{kes(r.unit.rent + r.unit.garbage)}</div>
                      <div className="text-[11px]" style={{ color: MUTED }}>Rent + services</div>
                    </td>
                    <td className="px-3 py-3 tabnum whitespace-nowrap">
                      <div className="font-semibold" style={{ color: r.balance > 0 ? "#C0392B" : INK }}>{kes(r.balance)}</div>
                      <div className="text-[11px]" style={{ color: MUTED }}>{r.unit.occupancyStatus === "vacant" ? "—" : r.balance <= 0 ? "Settled" : `${r.monthsOut.toFixed(2)} months due`}</div>
                    </td>
                    <td className="px-3 py-3"><Pill label={r.priority.key === "HIGH" ? "In Arrears" : r.priority.key === "PAID" ? "Up To Date" : r.priority.key === "VACANT" ? "Vacant" : "Review"} colorKey={r.priority.key} /></td>
                    <td className="px-3 py-3 whitespace-nowrap" style={{ color: MUTED }}>{r.priority.action}</td>
                    <td className="px-3 py-3 whitespace-nowrap text-right"><span className="text-xs font-semibold inline-flex items-center gap-0.5" style={{ color: "#1F8A4C" }}>Open <ChevronRight size={13} /></span></td>
                  </tr>
                );
              })}
              {sorted.length === 0 && <tr><td colSpan={7} className="px-3 py-10 text-center text-sm" style={{ color: MUTED }}>No units match these filters.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
      <div className="rounded-2xl border mt-4 px-4 py-3 flex flex-wrap items-center justify-between gap-2 bg-white" style={{ borderColor: BORDER }}>
        <div className="text-xs" style={{ color: MUTED }}><span className="font-semibold" style={{ color: INK }}>Tip:</span> Open a unit to add this month's charge, record an M-Pesa payment, or print a statement.</div>
        <button onClick={goImportExport} className="text-xs font-semibold shrink-0" style={{ color: "#1F8A4C" }}>Import history</button>
      </div>
      {!canEdit && <p className="text-xs mt-3" style={{ color: MUTED }}>Viewing as Read-only — actions are hidden.</p>}
      {addUnitOpen && (
        <AddUnitModal
          data={data}
          onClose={() => setAddUnitOpen(false)}
          onSubmit={(payload) => {
            const res = onAddUnit(payload);
            if (res && res.unitId) { setAddUnitOpen(false); openUnit(res.unitId); }
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Unit detail                                                              */
/* ---------------------------------------------------------------------- */
function UnitDetail({ data, unitId, canEdit, back, onStatement, onCloseTenancy, recordPayment, addMonthlyCharge, addAdjustment, updateUnitCharges, addTenantToUnit, updateTenant, updateTransaction, deleteTransaction }) {
  const unit = data.units.find((u) => u.id === unitId);
  const tenant = data.tenants.find((t) => t.id === unit.currentTenantId);
  const txns = data.transactions.filter((t) => t.unitId === unitId);
  const ledger = ledgerWithRunningBalance(txns).reverse();
  const balance = ledger.length ? ledger[0].runningBalance : 0;
  const status = statusOf(unit, balance);
  const priority = priorityOf(status);

  const [payOpen, setPayOpen] = useState(false);
  const [chargeOpen, setChargeOpen] = useState(false);
  const [adjOpen, setAdjOpen] = useState(false);
  const [chargesOpen, setChargesOpen] = useState(false);
  const [addTenantOpen, setAddTenantOpen] = useState(false);
  const [editTenantOpen, setEditTenantOpen] = useState(false);
  const [chargeMonth, setChargeMonth] = useState(todayISO().slice(0, 7));
  const [dupWarn, setDupWarn] = useState(null);
  const [editTxn, setEditTxn] = useState(null); // the payment/charge currently being edited
  const [deleteTxn, setDeleteTxn] = useState(null); // the payment/charge pending delete confirmation

  return (
    <div className="pt-2">
      <button onClick={back} className="flex items-center gap-1 text-sm font-semibold mb-3" style={{ color: MUTED }}><ArrowLeft size={15} /> Units &amp; balances</button>

      <div className="rounded-2xl border bg-white p-5 mb-5" style={{ borderColor: BORDER }}>
        <div className="flex flex-wrap items-start justify-between gap-3 mb-3">
          <div>
            <div className="text-xs font-semibold" style={{ color: MUTED }}>{unit.type === "shop" ? "Shop" : "Residential"}</div>
            <h1 className="text-xl font-bold">{unit.code}</h1>
            <div className="mt-1">{tenant ? <span className="text-sm font-semibold">{tenant.name}</span> : <span className="text-sm" style={{ color: MUTED }}>Vacant</span>}</div>
            {tenant && <div className="text-xs mt-0.5" style={{ color: MUTED }}>{tenant.phone} {tenant.email ? `· ${tenant.email}` : ""} · Moved in {fmtDate(tenant.moveInDate)}</div>}
          </div>
          <Pill label={priority.key === "HIGH" ? "In Arrears" : priority.key === "PAID" ? "Up To Date" : priority.key === "VACANT" ? "Vacant" : "Review"} colorKey={priority.key} />
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-4">
          <div className="rounded-xl p-3" style={{ background: BEIGE }}><div className="text-[11px] font-semibold mb-0.5" style={{ color: MUTED }}>Rent</div><div className="tabnum text-sm font-semibold">{kes(unit.rent)}</div></div>
          <div className="rounded-xl p-3" style={{ background: BEIGE }}><div className="text-[11px] font-semibold mb-0.5" style={{ color: MUTED }}>Garbage</div><div className="tabnum text-sm font-semibold">{kes(unit.garbage)}</div></div>
          <div className="rounded-xl p-3" style={{ background: BEIGE }}><div className="text-[11px] font-semibold mb-0.5" style={{ color: MUTED }}>Deposit Held</div><div className="tabnum text-sm font-semibold">{kes(unit.occupancyStatus === "vacant" ? 0 : unit.depositHeld)}</div></div>
          <div className="rounded-xl p-3" style={{ background: BEIGE }}><div className="text-[11px] font-semibold mb-0.5" style={{ color: MUTED }}>Current Balance</div><div className="tabnum text-base font-bold">{kes(unit.occupancyStatus === "vacant" ? 0 : balance)}</div></div>
        </div>

        {canEdit ? (
          <div className="flex flex-wrap gap-2">
            <Btn onClick={() => setPayOpen(true)} disabled={unit.occupancyStatus === "vacant"}>Record Payment</Btn>
            <Btn variant="secondary" onClick={() => { setDupWarn(null); setChargeOpen(true); }} disabled={unit.occupancyStatus === "vacant"}>Add Monthly Charge</Btn>
            <Btn variant="secondary" onClick={() => setAdjOpen(true)} disabled={unit.occupancyStatus === "vacant"}>Manual Adjustment</Btn>
            <Btn variant="secondary" onClick={() => setChargesOpen(true)}>Edit Charges</Btn>
            {unit.occupancyStatus === "vacant" ? (
              <Btn variant="secondary" onClick={() => setAddTenantOpen(true)}>Add Tenant</Btn>
            ) : (
              <Btn variant="secondary" onClick={() => setEditTenantOpen(true)} disabled={!tenant}>Edit Tenant</Btn>
            )}
            <Btn variant="secondary" onClick={onStatement} disabled={!tenant}>View Statement</Btn>
            <Btn variant="danger" onClick={onCloseTenancy} disabled={!tenant}>Close Tenancy</Btn>
          </div>
        ) : (
          <Btn variant="secondary" onClick={onStatement} disabled={!tenant}>View Statement</Btn>
        )}
      </div>

      <div className="rounded-2xl border bg-white overflow-hidden" style={{ borderColor: BORDER }}>
        <div className="px-4 py-3 border-b font-bold text-sm" style={{ borderColor: BORDER }}>Ledger</div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead style={{ borderBottom: `1px solid ${BORDER}` }}>
              <tr>
                {["Date", "Month", "Description", "Rent Charged", "Garbage Charged", "Payment Received", "M-Pesa Ref", "Running Balance", "Remarks"].map((h) => <th key={h} className="text-left px-3 py-2 text-xs font-bold whitespace-nowrap" style={{ color: MUTED }}>{h}</th>)}
                <th className="text-left px-2 py-2 text-xs font-bold whitespace-nowrap sticky right-0 z-10" style={{ color: MUTED, background: CARD, boxShadow: "-6px 0 6px -6px rgba(0,0,0,0.12)" }}></th>
              </tr>
            </thead>
            <tbody>
              {ledger.map((t) => {
                const editable = canEdit && (t.type === "payment" || t.type === "charge");
                return (
                  <tr key={t.id} style={{ borderBottom: `1px solid ${BORDER}` }}>
                    <td className="px-3 py-2 whitespace-nowrap">{fmtDate(t.date)}</td>
                    <td className="px-3 py-2 whitespace-nowrap" style={{ color: MUTED }}>{monthLabel(t.month)}</td>
                    <td className="px-3 py-2 whitespace-nowrap">{t.description}</td>
                    <td className="px-3 py-2 tabnum whitespace-nowrap">{t.rentCharged ? kes(t.rentCharged) : "—"}</td>
                    <td className="px-3 py-2 tabnum whitespace-nowrap">{t.garbageCharged ? kes(t.garbageCharged) : "—"}</td>
                    <td className="px-3 py-2 tabnum whitespace-nowrap" style={{ color: t.paymentReceived ? "#1F8A4C" : INK }}>{t.paymentReceived ? kes(t.paymentReceived) : "—"}</td>
                    <td className="px-3 py-2 whitespace-nowrap font-mono text-xs">{t.mpesaRef || "—"}</td>
                    <td className="px-3 py-2 tabnum whitespace-nowrap font-semibold">{kes(t.runningBalance)}</td>
                    <td className="px-3 py-2 whitespace-nowrap" style={{ color: MUTED }}>{t.remarks || "—"}</td>
                    <td className="px-2 py-2 whitespace-nowrap text-right sticky right-0 z-10" style={{ background: CARD, boxShadow: "-6px 0 6px -6px rgba(0,0,0,0.12)" }}>
                      {editable && <RowActionsMenu onEdit={() => setEditTxn(t)} onDelete={() => setDeleteTxn(t)} />}
                    </td>
                  </tr>
                );
              })}
              {ledger.length === 0 && <tr><td colSpan={10} className="px-3 py-8 text-center text-sm" style={{ color: MUTED }}>No transactions yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {payOpen && <PayModal unit={unit} onClose={() => setPayOpen(false)} onSubmit={(p) => { recordPayment(unit, p); setPayOpen(false); }} />}
      {chargeOpen && (
        <Modal title="Add Monthly Charge" onClose={() => setChargeOpen(false)}>
          <Field label="Month"><input type="month" value={chargeMonth} onChange={(e) => { setChargeMonth(e.target.value); setDupWarn(null); }} className={inputCls} style={inputStyle} /></Field>
          <p className="text-xs mb-3" style={{ color: MUTED }}>Will charge {kes(unit.rent)} rent + {kes(unit.garbage)} garbage, dated the 1st of the month.</p>
          {dupWarn && <div className="flex items-start gap-2 text-xs bg-amber-50 text-amber-800 rounded-lg p-2.5 mb-3"><AlertTriangle size={14} className="mt-0.5 shrink-0" />A charge for {monthLabel(chargeMonth)} already exists on this unit. Adding another will double-charge the tenant.</div>}
          <div className="flex justify-end gap-2">
            <Btn variant="secondary" onClick={() => setChargeOpen(false)}>Cancel</Btn>
            <Btn onClick={() => { const { dup, commit } = addMonthlyCharge(unit, `${chargeMonth}-01`); if (dup && !dupWarn) { setDupWarn(true); return; } commit(); setChargeOpen(false); setDupWarn(null); }}>{dupWarn ? "Add Anyway" : "Add Charge"}</Btn>
          </div>
        </Modal>
      )}
      {adjOpen && <AdjustModal unit={unit} onClose={() => setAdjOpen(false)} onSubmit={(p) => { addAdjustment(unit, p); setAdjOpen(false); }} />}
      {chargesOpen && (
        <EditChargesModal unit={unit} onClose={() => setChargesOpen(false)} onSubmit={(r, g) => { updateUnitCharges(unit.id, r, g); setChargesOpen(false); }} />
      )}
      {addTenantOpen && (
        <AddTenantModal unit={unit} onClose={() => setAddTenantOpen(false)} onSubmit={(payload) => { addTenantToUnit(unit.id, payload); setAddTenantOpen(false); }} />
      )}
      {editTenantOpen && tenant && (
        <EditTenantModal tenant={tenant} onClose={() => setEditTenantOpen(false)} onSubmit={(payload) => { updateTenant(tenant.id, payload); setEditTenantOpen(false); }} />
      )}

      {editTxn && editTxn.type === "payment" && (
        <EditPaymentModal
          unit={unit} txn={editTxn} onClose={() => setEditTxn(null)}
          onSubmit={(payload) => { if (updateTransaction(editTxn.id, payload)) setEditTxn(null); }}
        />
      )}
      {editTxn && editTxn.type === "charge" && (
        <EditChargeModal
          unit={unit} txn={editTxn} onClose={() => setEditTxn(null)}
          onSubmit={(payload) => { if (updateTransaction(editTxn.id, payload)) setEditTxn(null); }}
        />
      )}
      {deleteTxn && (
        <DeleteTransactionModal
          unit={unit} txn={deleteTxn} onClose={() => setDeleteTxn(null)}
          onConfirm={() => { deleteTransaction(deleteTxn.id); setDeleteTxn(null); }}
        />
      )}
    </div>
  );
}

function EditChargesModal({ unit, onClose, onSubmit }) {
  const [rent, setRent] = useState(String(unit.rent ?? 0));
  const [garbage, setGarbage] = useState(String(unit.garbage ?? 0));
  const [error, setError] = useState("");
  const total = (Number(rent) || 0) + (Number(garbage) || 0);

  function handleSave() {
    const r = Number(rent), g = Number(garbage);
    if (rent === "" || isNaN(r) || r < 0) { setError("Monthly Rent cannot be negative."); return; }
    if (garbage === "" || isNaN(g) || g < 0) { setError("Monthly Garbage cannot be negative."); return; }
    onSubmit(r, g);
  }

  return (
    <Modal title={`Edit Unit Charges — ${unit.code}`} onClose={onClose}>
      <Field label="Monthly Rent (KES)"><input type="number" min="0" value={rent} onChange={(e) => { setRent(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
      <Field label="Monthly Garbage (KES)"><input type="number" min="0" value={garbage} onChange={(e) => { setGarbage(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
      <div className="rounded-xl p-3 mb-3" style={{ background: BEIGE }}>
        <div className="text-[11px] font-semibold mb-0.5" style={{ color: MUTED }}>Total Monthly Charge</div>
        <div className="tabnum text-base font-bold">{kes(total)}</div>
      </div>
      <p className="text-xs mb-3" style={{ color: MUTED }}>This only changes future monthly charges. Past ledger entries are never rewritten.</p>
      {error && <div className="text-xs text-red-700 bg-red-50 rounded-lg p-2.5 mb-3">{error}</div>}
      <div className="flex justify-end gap-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={handleSave}>Save Changes</Btn>
      </div>
    </Modal>
  );
}

function AddTenantModal({ unit, onClose, onSubmit }) {
  const [name, setName] = useState("");
  const [phone, setPhone] = useState("");
  const [email, setEmail] = useState("");
  const [moveInDate, setMoveInDate] = useState(todayISO());
  const [notes, setNotes] = useState("");
  const [rent, setRent] = useState(String(unit.rent ?? ""));
  const [garbage, setGarbage] = useState(String(unit.garbage ?? ""));
  const [deposit, setDeposit] = useState("0");
  const [createFirstCharge, setCreateFirstCharge] = useState(false);
  const [error, setError] = useState("");

  function handleSave() {
    if (!name.trim()) { setError("Tenant name is required."); return; }
    if (!moveInDate) { setError("Move-in date is required."); return; }
    const r = Number(rent), g = Number(garbage);
    if (rent === "" || isNaN(r) || r < 0) { setError("Monthly Rent is required and cannot be negative."); return; }
    if (garbage === "" || isNaN(g) || g < 0) { setError("Monthly Garbage is required and cannot be negative."); return; }
    const dep = deposit === "" ? 0 : Number(deposit);
    if (isNaN(dep) || dep < 0) { setError("Deposit cannot be negative."); return; }
    onSubmit({ name, phone, email, moveInDate, notes, rent: r, garbage: g, deposit: dep, createFirstCharge });
  }

  return (
    <Modal title="Add Tenant" onClose={onClose} wide>
      <Field label="Unit"><div className="text-sm font-semibold px-3 py-2 rounded-xl" style={{ background: BEIGE }}>{unit.code}</div></Field>

      <div className="text-xs font-bold tracking-wide mt-2 mb-1" style={{ color: MUTED }}>TENANT INFORMATION</div>
      <Field label="Full Name"><input value={name} onChange={(e) => { setName(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Phone Number"><input value={phone} onChange={(e) => setPhone(e.target.value)} className={inputCls} style={inputStyle} /></Field>
        <Field label="Email"><input value={email} onChange={(e) => setEmail(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      </div>
      <Field label="Move-in Date"><input type="date" value={moveInDate} onChange={(e) => { setMoveInDate(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
      <Field label="Notes"><textarea value={notes} onChange={(e) => setNotes(e.target.value)} className={inputCls} style={inputStyle} rows={2} /></Field>

      <div className="text-xs font-bold tracking-wide mt-3 mb-1" style={{ color: MUTED }}>CHARGES</div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Monthly Rent (KES)"><input type="number" min="0" value={rent} onChange={(e) => { setRent(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
        <Field label="Monthly Garbage (KES)"><input type="number" min="0" value={garbage} onChange={(e) => { setGarbage(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
      </div>

      <div className="text-xs font-bold tracking-wide mt-3 mb-1" style={{ color: MUTED }}>DEPOSIT</div>
      <Field label="Deposit Received (KES)"><input type="number" min="0" value={deposit} onChange={(e) => { setDeposit(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
      <label className="flex items-center gap-2 text-xs font-semibold mb-3"><input type="checkbox" checked={createFirstCharge} onChange={(e) => setCreateFirstCharge(e.target.checked)} /> Create first month's rent charge now</label>

      {error && <div className="text-xs text-red-700 bg-red-50 rounded-lg p-2.5 mb-3">{error}</div>}
      <div className="flex justify-end gap-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={handleSave}>Save</Btn>
      </div>
    </Modal>
  );
}

function EditTenantModal({ tenant, onClose, onSubmit }) {
  const [name, setName] = useState(tenant.name || "");
  const [phone, setPhone] = useState(tenant.phone || "");
  const [email, setEmail] = useState(tenant.email || "");
  const [moveInDate, setMoveInDate] = useState(tenant.moveInDate || "");
  const [notes, setNotes] = useState(tenant.notes || "");
  const [error, setError] = useState("");

  function handleSave() {
    if (!name.trim()) { setError("Tenant name is required."); return; }
    if (!moveInDate) { setError("Move-in date is required."); return; }
    onSubmit({ name, phone, email, moveInDate, notes });
  }

  return (
    <Modal title="Edit Tenant" onClose={onClose}>
      <Field label="Full Name"><input value={name} onChange={(e) => { setName(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
      <Field label="Phone Number"><input value={phone} onChange={(e) => setPhone(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Email"><input value={email} onChange={(e) => setEmail(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Move-in Date"><input type="date" value={moveInDate} onChange={(e) => { setMoveInDate(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
      <Field label="Notes"><textarea value={notes} onChange={(e) => setNotes(e.target.value)} className={inputCls} style={inputStyle} rows={2} /></Field>
      {error && <div className="text-xs text-red-700 bg-red-50 rounded-lg p-2.5 mb-3">{error}</div>}
      <div className="flex justify-end gap-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={handleSave}>Save Changes</Btn>
      </div>
    </Modal>
  );
}

function AddUnitModal({ data, onClose, onSubmit }) {
  const [code, setCode] = useState("");
  const [type, setType] = useState("residential");
  const [occupancy, setOccupancy] = useState("vacant");
  const [rent, setRent] = useState("0");
  const [garbage, setGarbage] = useState("0");
  const [tenantName, setTenantName] = useState("");
  const [phone, setPhone] = useState("");
  const [email, setEmail] = useState("");
  const [moveInDate, setMoveInDate] = useState(todayISO());
  const [notes, setNotes] = useState("");
  const [deposit, setDeposit] = useState("0");
  const [createFirstCharge, setCreateFirstCharge] = useState(false);
  const [error, setError] = useState("");

  function handleSave() {
    const trimmed = code.trim();
    if (!trimmed) { setError("Unit Code is required."); return; }
    const exists = data.units.some((u) => u.code.trim().toLowerCase() === trimmed.toLowerCase());
    if (exists) { setError(`A unit with code ${trimmed} already exists.`); return; }
    const r = rent === "" ? 0 : Number(rent);
    const g = garbage === "" ? 0 : Number(garbage);
    if (isNaN(r) || r < 0) { setError("Rent cannot be negative."); return; }
    if (isNaN(g) || g < 0) { setError("Garbage cannot be negative."); return; }

    if (occupancy === "tenant") {
      if (!tenantName.trim()) { setError("Tenant name is required."); return; }
      if (!moveInDate) { setError("Move-in date is required."); return; }
      const dep = deposit === "" ? 0 : Number(deposit);
      if (isNaN(dep) || dep < 0) { setError("Deposit cannot be negative."); return; }
      onSubmit({ code: trimmed, type, rent: r, garbage: g, addTenantNow: true, tenantName, phone, email, moveInDate, notes, deposit: dep, createFirstCharge });
    } else {
      onSubmit({ code: trimmed, type, rent: r, garbage: g, addTenantNow: false });
    }
  }

  return (
    <Modal title="Add Unit" onClose={onClose} wide>
      <div className="text-xs font-bold tracking-wide mb-1" style={{ color: MUTED }}>UNIT INFORMATION</div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Unit Code"><input value={code} onChange={(e) => { setCode(e.target.value); setError(""); }} className={inputCls} style={inputStyle} placeholder="e.g. D1" /></Field>
        <Field label="Unit Type">
          <select value={type} onChange={(e) => setType(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="residential">Residential</option>
            <option value="shop">Shop</option>
          </select>
        </Field>
      </div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Monthly Rent (KES)"><input type="number" min="0" value={rent} onChange={(e) => { setRent(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
        <Field label="Monthly Garbage (KES)"><input type="number" min="0" value={garbage} onChange={(e) => { setGarbage(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
      </div>

      <div className="text-xs font-bold tracking-wide mt-3 mb-1" style={{ color: MUTED }}>OCCUPANCY</div>
      <div className="flex gap-2 mb-3">
        <button type="button" onClick={() => setOccupancy("vacant")} className="flex-1 rounded-xl border px-3 py-2 text-sm font-semibold"
          style={{ borderColor: occupancy === "vacant" ? "#1F8A4C" : BORDER, background: occupancy === "vacant" ? "#EDF8F1" : "white", color: occupancy === "vacant" ? "#1F8A4C" : INK }}>
          Vacant
        </button>
        <button type="button" onClick={() => setOccupancy("tenant")} className="flex-1 rounded-xl border px-3 py-2 text-sm font-semibold"
          style={{ borderColor: occupancy === "tenant" ? "#1F8A4C" : BORDER, background: occupancy === "tenant" ? "#EDF8F1" : "white", color: occupancy === "tenant" ? "#1F8A4C" : INK }}>
          Add New Tenant Now
        </button>
      </div>

      {occupancy === "tenant" && (
        <>
          <div className="text-xs font-bold tracking-wide mt-2 mb-1" style={{ color: MUTED }}>TENANT</div>
          <Field label="Full Name"><input value={tenantName} onChange={(e) => { setTenantName(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Phone"><input value={phone} onChange={(e) => setPhone(e.target.value)} className={inputCls} style={inputStyle} /></Field>
            <Field label="Email"><input value={email} onChange={(e) => setEmail(e.target.value)} className={inputCls} style={inputStyle} /></Field>
          </div>
          <Field label="Move-in Date"><input type="date" value={moveInDate} onChange={(e) => { setMoveInDate(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
          <Field label="Notes"><textarea value={notes} onChange={(e) => setNotes(e.target.value)} className={inputCls} style={inputStyle} rows={2} /></Field>

          <div className="text-xs font-bold tracking-wide mt-2 mb-1" style={{ color: MUTED }}>DEPOSIT</div>
          <Field label="Deposit Received (KES)"><input type="number" min="0" value={deposit} onChange={(e) => { setDeposit(e.target.value); setError(""); }} className={inputCls} style={inputStyle} /></Field>
          <label className="flex items-center gap-2 text-xs font-semibold mb-3"><input type="checkbox" checked={createFirstCharge} onChange={(e) => setCreateFirstCharge(e.target.checked)} /> Create first month's rent charge now</label>
        </>
      )}

      {error && <div className="text-xs text-red-700 bg-red-50 rounded-lg p-2.5 mb-3">{error}</div>}
      <div className="flex justify-end gap-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={handleSave}>Save</Btn>
      </div>
    </Modal>
  );
}

function PayModal({ unit, onClose, onSubmit }) {
  const [date, setDate] = useState(todayISO());
  const [amount, setAmount] = useState("");
  const [mpesaRef, setMpesaRef] = useState("");
  const [method, setMethod] = useState("M-Pesa");
  const [remarks, setRemarks] = useState("");
  return (
    <Modal title={`Record Payment — ${unit.code}`} onClose={onClose}>
      <Field label="Date"><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Amount (KES)"><input type="number" value={amount} onChange={(e) => setAmount(e.target.value)} className={inputCls} style={inputStyle} placeholder="0" /></Field>
      <Field label="Payment method"><select value={method} onChange={(e) => setMethod(e.target.value)} className={inputCls} style={inputStyle}><option>M-Pesa</option><option>Bank Transfer</option><option>Cash</option><option>Other</option></select></Field>
      <Field label="M-Pesa Reference" hint="Leave blank if paid by cash or bank"><input value={mpesaRef} onChange={(e) => setMpesaRef(e.target.value)} className={inputCls} style={inputStyle} placeholder="e.g. QFT7X9K2L1" /></Field>
      <Field label="Remarks (optional)"><input value={remarks} onChange={(e) => setRemarks(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn disabled={!amount} onClick={() => onSubmit({ date, amount, mpesaRef, method, remarks })}>Record Payment</Btn>
      </div>
    </Modal>
  );
}
function EditPaymentModal({ unit, txn, onClose, onSubmit }) {
  const [date, setDate] = useState(txn.date || todayISO());
  const [amount, setAmount] = useState(String(txn.paymentReceived ?? ""));
  const [mpesaRef, setMpesaRef] = useState(txn.mpesaRef || "");
  const [description, setDescription] = useState(txn.description || "");
  const [remarks, setRemarks] = useState(txn.remarks || "");
  return (
    <Modal title={`Edit Payment — ${unit.code}`} onClose={onClose}>
      <Field label="Date"><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Amount (KES)"><input type="number" value={amount} onChange={(e) => setAmount(e.target.value)} className={inputCls} style={inputStyle} placeholder="0" /></Field>
      <Field label="M-Pesa Reference" hint="Leave blank if paid by cash or bank"><input value={mpesaRef} onChange={(e) => setMpesaRef(e.target.value)} className={inputCls} style={inputStyle} placeholder="e.g. QFT7X9K2L1" /></Field>
      <Field label="Description"><input value={description} onChange={(e) => setDescription(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Remarks"><input value={remarks} onChange={(e) => setRemarks(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn disabled={!amount} onClick={() => onSubmit({ date, paymentReceived: amount, mpesaRef, description, remarks })}>Save Changes</Btn>
      </div>
    </Modal>
  );
}

function EditChargeModal({ unit, txn, onClose, onSubmit }) {
  const [date, setDate] = useState(txn.date || todayISO());
  const [rentCharged, setRentCharged] = useState(String(txn.rentCharged ?? 0));
  const [garbageCharged, setGarbageCharged] = useState(String(txn.garbageCharged ?? 0));
  const [description, setDescription] = useState(txn.description || "");
  const [remarks, setRemarks] = useState(txn.remarks || "");
  return (
    <Modal title={`Edit Monthly Charge — ${unit.code}`} onClose={onClose}>
      <Field label="Charge Date"><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Rent Amount (KES)"><input type="number" min="0" value={rentCharged} onChange={(e) => setRentCharged(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Garbage Amount (KES)"><input type="number" min="0" value={garbageCharged} onChange={(e) => setGarbageCharged(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Description"><input value={description} onChange={(e) => setDescription(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Remarks"><input value={remarks} onChange={(e) => setRemarks(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => onSubmit({ date, rentCharged, garbageCharged, description, remarks })}>Save Changes</Btn>
      </div>
    </Modal>
  );
}

// Shared confirmation dialog for deleting either a Payment or a Monthly
// Charge — wording adapts per type, but the user always has the final say:
// confirmation guards against accidental clicks, it never blocks a
// deliberate delete.
function DeleteTransactionModal({ unit, txn, onClose, onConfirm }) {
  const isPayment = txn.type === "payment";
  const amount = isPayment ? Number(txn.paymentReceived) || 0 : (Number(txn.rentCharged) || 0) + (Number(txn.garbageCharged) || 0);
  return (
    <Modal title={isPayment ? "Delete Payment" : "Delete Monthly Charge"} onClose={onClose}>
      <div className="rounded-xl p-3 mb-4" style={{ background: BEIGE }}>
        <div className="text-[11px] font-semibold mb-1" style={{ color: MUTED }}>{unit.code} · {fmtDate(txn.date)}</div>
        <div className="text-sm font-semibold">{txn.description}</div>
        <div className="tabnum text-base font-bold mt-1">{kes(amount)}</div>
        {txn.mpesaRef && <div className="text-xs font-mono mt-0.5" style={{ color: MUTED }}>{txn.mpesaRef}</div>}
      </div>
      <p className="text-sm mb-5">
        {isPayment
          ? "Are you sure you want to delete this payment? This action cannot be undone."
          : "Are you sure you want to delete this monthly charge? Deleting it will affect the unit balance and monthly totals."}
      </p>
      <div className="flex justify-end gap-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn variant="danger" icon={Trash2} onClick={onConfirm}>{isPayment ? "Delete Payment" : "Delete Monthly Charge"}</Btn>
      </div>
    </Modal>
  );
}

function AdjustModal({ unit, onClose, onSubmit }) {
  const [date, setDate] = useState(todayISO());
  const [description, setDescription] = useState("Damage charge");
  const [charge, setCharge] = useState("");
  const [credit, setCredit] = useState("");
  const [remarks, setRemarks] = useState("");
  return (
    <Modal title={`Manual Adjustment — ${unit.code}`} onClose={onClose}>
      <Field label="Date"><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Description"><input value={description} onChange={(e) => setDescription(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <Field label="Charge to add (KES)" hint="Increases balance — e.g. damage, penalty"><input type="number" value={charge} onChange={(e) => setCharge(e.target.value)} className={inputCls} style={inputStyle} placeholder="0" /></Field>
      <Field label="Credit to add (KES)" hint="Decreases balance — e.g. goodwill credit, correction"><input type="number" value={credit} onChange={(e) => setCredit(e.target.value)} className={inputCls} style={inputStyle} placeholder="0" /></Field>
      <Field label="Remarks"><input value={remarks} onChange={(e) => setRemarks(e.target.value)} className={inputCls} style={inputStyle} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => onSubmit({ date, description, charge, credit, remarks })}>Save Adjustment</Btn>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/* Statement                                                                 */
/* ---------------------------------------------------------------------- */
function Statement({ data, unitId, back }) {
  const unit = data.units.find((u) => u.id === unitId);
  const tenant = data.tenants.find((t) => t.id === unit.currentTenantId);
  const allTxns = data.transactions.filter((t) => t.unitId === unitId);
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const filtered = allTxns.filter((t) => (!from || t.date >= from) && (!to || t.date <= to));
  const ledger = ledgerWithRunningBalance(filtered);
  const fullLedger = ledgerWithRunningBalance(allTxns);
  const balance = fullLedger.length ? fullLedger[fullLedger.length - 1].runningBalance : 0;

  return (
    <div className="pt-2">
      <div className="flex items-center justify-between mb-3 print:hidden">
        <button onClick={back} className="flex items-center gap-1 text-sm font-semibold" style={{ color: MUTED }}><ArrowLeft size={15} /> Back</button>
        <Btn icon={Printer} onClick={() => window.print()}>Print / Save PDF</Btn>
      </div>
      <div className="rounded-2xl border bg-white p-6" style={{ borderColor: BORDER }}>
        <div className="flex items-center justify-between mb-4">
          <div>
            <div className="text-xs font-bold uppercase tracking-wide" style={{ color: MUTED }}>Tenant Statement</div>
            <h1 className="text-lg font-bold">{unit.code}</h1>
          </div>
          <div style={{ background: SIDEBAR }} className="text-white text-xs font-bold px-3 py-1.5 rounded-lg">MaliDesk</div>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-4 text-sm">
          <div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Tenant</div><div className="font-semibold">{tenant?.name}</div></div>
          <div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Monthly Rent</div><div className="tabnum">{kes(unit.rent)}</div></div>
          <div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Garbage Fee</div><div className="tabnum">{kes(unit.garbage)}</div></div>
          <div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Deposit Held</div><div className="tabnum">{kes(unit.depositHeld)}</div></div>
        </div>
        <div className="mb-4 text-sm"><span className="text-[11px] font-semibold" style={{ color: MUTED }}>Current Balance </span><span className="font-bold tabnum">{kes(balance)}</span></div>
        <div className="flex items-center gap-2 mb-3 print:hidden">
          <Field label="From"><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} className={inputCls} style={inputStyle} /></Field>
          <Field label="To"><input type="date" value={to} onChange={(e) => setTo(e.target.value)} className={inputCls} style={inputStyle} /></Field>
        </div>
        <table className="w-full text-sm">
          <thead style={{ borderBottom: `1px solid ${BORDER}` }}><tr>{["Date", "Description", "Rent", "Garbage", "Payment", "M-Pesa Ref", "Balance"].map((h) => <th key={h} className="text-left px-2 py-2 text-xs font-bold" style={{ color: MUTED }}>{h}</th>)}</tr></thead>
          <tbody>
            {ledger.map((t) => (
              <tr key={t.id} style={{ borderBottom: `1px solid ${BORDER}` }}>
                <td className="px-2 py-1.5 whitespace-nowrap">{fmtDate(t.date)}</td><td className="px-2 py-1.5">{t.description}</td>
                <td className="px-2 py-1.5 tabnum">{t.rentCharged ? kes(t.rentCharged) : "—"}</td><td className="px-2 py-1.5 tabnum">{t.garbageCharged ? kes(t.garbageCharged) : "—"}</td>
                <td className="px-2 py-1.5 tabnum">{t.paymentReceived ? kes(t.paymentReceived) : "—"}</td><td className="px-2 py-1.5 font-mono text-xs">{t.mpesaRef || "—"}</td>
                <td className="px-2 py-1.5 tabnum font-semibold">{kes(t.runningBalance)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Close tenancy                                                            */
/* ---------------------------------------------------------------------- */
function CloseTenancyFlow({ data, unitId, back, onSubmit }) {
  const unit = data.units.find((u) => u.id === unitId);
  const tenant = data.tenants.find((t) => t.id === unit.currentTenantId);
  const txns = data.transactions.filter((t) => t.unitId === unitId);
  const finalBalance = unitBalance(txns);

  const [moveOutDate, setMoveOutDate] = useState(todayISO());
  const [reasonForLeaving, setReasonForLeaving] = useState("");
  const [arrearsOutcome, setArrearsOutcome] = useState("Not applicable");
  const [amountRefunded, setAmountRefunded] = useState(finalBalance <= 0 ? unit.depositHeld : 0);
  const [amountForfeited, setAmountForfeited] = useState(finalBalance > 0 ? Math.min(finalBalance, unit.depositHeld) : 0);
  const [isEviction, setIsEviction] = useState(false);
  const [evictionNoticeDate, setEvictionNoticeDate] = useState("");
  const [evictionLockDate, setEvictionLockDate] = useState("");
  const [recoveryNotes, setRecoveryNotes] = useState("");
  const [notes, setNotes] = useState("");
  const [confirming, setConfirming] = useState(false);

  return (
    <div className="pt-2">
      <button onClick={back} className="flex items-center gap-1 text-sm font-semibold mb-3" style={{ color: MUTED }}><ArrowLeft size={15} /> Back</button>
      <div className="rounded-2xl border bg-white p-5 max-w-2xl" style={{ borderColor: BORDER }}>
        <h1 className="text-lg font-bold mb-1">Close Tenancy — {unit.code}</h1>
        <p className="text-sm mb-4" style={{ color: MUTED }}>{tenant?.name} will be moved to the permanent archive. The unit becomes vacant with a clean balance.</p>
        <div className="rounded-xl p-3 mb-4 flex flex-wrap gap-6" style={{ background: BEIGE }}>
          <div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Final Balance</div><div className="font-bold tabnum">{kes(finalBalance)}</div></div>
          <div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Deposit Held</div><div className="font-bold tabnum">{kes(unit.depositHeld)}</div></div>
        </div>
        <Field label="Move-out date"><input type="date" value={moveOutDate} onChange={(e) => setMoveOutDate(e.target.value)} className={inputCls} style={inputStyle} /></Field>
        <Field label="Reason for leaving"><input value={reasonForLeaving} onChange={(e) => setReasonForLeaving(e.target.value)} className={inputCls} style={inputStyle} placeholder="e.g. Relocated, end of lease…" /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Amount refunded (KES)"><input type="number" value={amountRefunded} onChange={(e) => setAmountRefunded(e.target.value)} className={inputCls} style={inputStyle} /></Field>
          <Field label="Amount forfeited / applied to arrears (KES)"><input type="number" value={amountForfeited} onChange={(e) => setAmountForfeited(e.target.value)} className={inputCls} style={inputStyle} /></Field>
        </div>
        <Field label="Arrears outcome"><select value={arrearsOutcome} onChange={(e) => setArrearsOutcome(e.target.value)} className={inputCls} style={inputStyle}>{["Settled in full", "Partial recovery", "Written off", "Still pursuing", "Not applicable"].map((o) => <option key={o}>{o}</option>)}</select></Field>
        <label className="flex items-center gap-2 text-sm font-semibold mb-3"><input type="checkbox" checked={isEviction} onChange={(e) => setIsEviction(e.target.checked)} /> This is an eviction</label>
        {isEviction && (
          <div className="rounded-xl p-3 mb-3 border" style={{ borderColor: BORDER }}>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Notice date"><input type="date" value={evictionNoticeDate} onChange={(e) => setEvictionNoticeDate(e.target.value)} className={inputCls} style={inputStyle} /></Field>
              <Field label="Eviction / lock date"><input type="date" value={evictionLockDate} onChange={(e) => setEvictionLockDate(e.target.value)} className={inputCls} style={inputStyle} /></Field>
            </div>
            <Field label="Recovery notes / case reference"><textarea value={recoveryNotes} onChange={(e) => setRecoveryNotes(e.target.value)} className={inputCls} style={inputStyle} rows={2} /></Field>
          </div>
        )}
        <Field label="Notes"><textarea value={notes} onChange={(e) => setNotes(e.target.value)} className={inputCls} style={inputStyle} rows={2} /></Field>
        <div className="flex justify-end gap-2 mt-3">
          <Btn variant="secondary" onClick={back}>Cancel</Btn>
          <Btn variant="danger" onClick={() => setConfirming(true)}>Close Tenancy</Btn>
        </div>
      </div>
      {confirming && (
        <Modal title="Confirm: close this tenancy?" onClose={() => setConfirming(false)}>
          <p className="text-sm mb-4">This will archive <strong>{tenant?.name}</strong>'s full history for unit <strong>{unit.code}</strong> and mark the unit vacant. This cannot be undone from here.</p>
          <div className="flex justify-end gap-2">
            <Btn variant="secondary" onClick={() => setConfirming(false)}>Go Back</Btn>
            <Btn variant="danger" onClick={() => onSubmit(unit, tenant, { moveOutDate, reasonForLeaving, arrearsOutcome, amountRefunded, amountForfeited, evictionNoticeDate: isEviction ? evictionNoticeDate : "", evictionLockDate: isEviction ? evictionLockDate : "", recoveryNotes: isEviction ? recoveryNotes : "", notes })}>Confirm &amp; Archive</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Closed tenancies (archive)                                               */
/* ---------------------------------------------------------------------- */
function ArchiveList({ data, openDetail }) {
  const [q, setQ] = useState("");
  const rows = data.archive.filter((a) => !q || `${a.unitCode} ${a.tenant.name} ${a.reasonForLeaving}`.toLowerCase().includes(q.toLowerCase()));
  return (
    <div className="pt-2">
      <h1 className="text-xl font-bold mb-1">Closed tenancies</h1>
      <p className="text-sm mb-4" style={{ color: MUTED }}>Every past tenancy, kept exactly as it closed.</p>
      <div className="relative mb-3 max-w-sm">
        <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: MUTED }} />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search by tenant, unit, reason…" className={`${inputCls} pl-8`} style={inputStyle} />
      </div>
      <div className="rounded-2xl border overflow-hidden bg-white" style={{ borderColor: BORDER }}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead style={{ borderBottom: `1px solid ${BORDER}` }}><tr>{["Unit", "Tenant", "Move-out", "Final Balance", "Deposit", "Refunded", "Reason", "Arrears Outcome"].map((h) => <th key={h} className="text-left px-3 py-2.5 text-xs font-bold whitespace-nowrap" style={{ color: MUTED }}>{h}</th>)}</tr></thead>
            <tbody>
              {rows.map((a) => (
                <tr key={a.id} onClick={() => openDetail(a.id)} className="cursor-pointer hover:bg-black/[0.015]" style={{ borderBottom: `1px solid ${BORDER}` }}>
                  <td className="px-3 py-2.5 font-semibold whitespace-nowrap">{a.unitCode}</td><td className="px-3 py-2.5 whitespace-nowrap">{a.tenant.name}</td>
                  <td className="px-3 py-2.5 whitespace-nowrap">{fmtDate(a.tenant.moveOutDate)}</td><td className="px-3 py-2.5 tabnum whitespace-nowrap">{kes(a.finalBalance)}</td>
                  <td className="px-3 py-2.5 tabnum whitespace-nowrap">{kes(a.depositHeld)}</td><td className="px-3 py-2.5 tabnum whitespace-nowrap">{kes(a.amountRefunded)}</td>
                  <td className="px-3 py-2.5 whitespace-nowrap">{a.reasonForLeaving || "—"}</td><td className="px-3 py-2.5 whitespace-nowrap">{a.arrearsOutcome}</td>
                </tr>
              ))}
              {rows.length === 0 && <tr><td colSpan={8} className="px-3 py-10 text-center text-sm" style={{ color: MUTED }}>No closed tenancies yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
function ArchiveDetail({ data, archiveId, back }) {
  const a = data.archive.find((x) => x.id === archiveId);
  const ledger = ledgerWithRunningBalance(a.transactions);
  return (
    <div className="pt-2">
      <button onClick={back} className="flex items-center gap-1 text-sm font-semibold mb-3" style={{ color: MUTED }}><ArrowLeft size={15} /> Closed tenancies</button>
      <div className="rounded-2xl border bg-white p-5" style={{ borderColor: BORDER }}>
        <h1 className="text-lg font-bold mb-1">{a.tenant.name} — {a.unitCode}</h1>
        <p className="text-sm mb-4" style={{ color: MUTED }}>{fmtDate(a.tenant.moveInDate)} → {fmtDate(a.tenant.moveOutDate)}</p>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-4">
          <div className="rounded-xl p-3" style={{ background: BEIGE }}><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Final Balance</div><div className="tabnum text-base font-bold">{kes(a.finalBalance)}</div></div>
          <div className="rounded-xl p-3" style={{ background: BEIGE }}><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Deposit Held</div><div className="tabnum text-sm font-semibold">{kes(a.depositHeld)}</div></div>
          <div className="rounded-xl p-3" style={{ background: BEIGE }}><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Refunded</div><div className="tabnum text-sm font-semibold">{kes(a.amountRefunded)}</div></div>
          <div className="rounded-xl p-3" style={{ background: BEIGE }}><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Forfeited</div><div className="tabnum text-sm font-semibold">{kes(a.amountForfeited)}</div></div>
        </div>
        <div className="grid grid-cols-2 gap-3 text-sm mb-4">
          <div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Reason for leaving</div><div>{a.reasonForLeaving || "—"}</div></div>
          <div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Arrears outcome</div><div>{a.arrearsOutcome}</div></div>
          {a.evictionLockDate && <><div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Eviction notice date</div><div>{fmtDate(a.evictionNoticeDate)}</div></div><div><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Lock date</div><div>{fmtDate(a.evictionLockDate)}</div></div></>}
          {a.notes && <div className="col-span-2"><div className="text-[11px] font-semibold" style={{ color: MUTED }}>Notes</div><div>{a.notes}</div></div>}
        </div>
        <div className="font-bold text-sm mb-2">Full transaction history (immutable)</div>
        <table className="w-full text-sm">
          <thead style={{ borderBottom: `1px solid ${BORDER}` }}><tr>{["Date", "Description", "Rent", "Garbage", "Payment", "M-Pesa Ref", "Balance"].map((h) => <th key={h} className="text-left px-2 py-2 text-xs font-bold" style={{ color: MUTED }}>{h}</th>)}</tr></thead>
          <tbody>
            {ledger.map((t) => (
              <tr key={t.id} style={{ borderBottom: `1px solid ${BORDER}` }}>
                <td className="px-2 py-1.5 whitespace-nowrap">{fmtDate(t.date)}</td><td className="px-2 py-1.5">{t.description}</td>
                <td className="px-2 py-1.5 tabnum">{t.rentCharged ? kes(t.rentCharged) : "—"}</td><td className="px-2 py-1.5 tabnum">{t.garbageCharged ? kes(t.garbageCharged) : "—"}</td>
                <td className="px-2 py-1.5 tabnum">{t.paymentReceived ? kes(t.paymentReceived) : "—"}</td><td className="px-2 py-1.5 font-mono text-xs">{t.mpesaRef || "—"}</td>
                <td className="px-2 py-1.5 tabnum font-semibold">{kes(t.runningBalance)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Import & export                                                          */
/* ---------------------------------------------------------------------- */
function ImportExport({ data, onImport, onExportExcel, onExportCSV, showColumns }) {
  const fileRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  const [include, setInclude] = useState({ units: true, transactions: true, archive: false });
  const [pendingFile, setPendingFile] = useState(null);

  // Importing now overwrites all current units, tenants and transactions,
  // so a file is staged here and only handed to onImport once the user
  // confirms — never on drop/select alone.
  const handleFiles = (files) => { if (files && files[0]) setPendingFile(files[0]); };
  const confirmImport = () => { if (pendingFile) onImport(pendingFile); setPendingFile(null); };

  const items = [
    ["units", "Current units", "Units, rent and occupancy", data.units.length],
    ["transactions", "Transactions", "Charges, payments and balances", data.transactions.length],
    ["archive", "Closed tenancies", "Move-outs and deposit outcomes", data.archive.length],
  ];
  const selectedCount = items.reduce((s, [k, , , count]) => s + (include[k] ? count : 0), 0);

  return (
    <div className="pt-2">
      <h1 className="text-xl font-bold mb-1">Import &amp; export</h1>
      <p className="text-sm mb-6" style={{ color: MUTED }}>Move your records in and out of MaliDesk with a simple spreadsheet workflow. Your source file stays on this device.</p>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5 mb-6">
        <div className="rounded-2xl border bg-white p-5" style={{ borderColor: BORDER }}>
          <div className="text-[11px] font-bold tracking-wide" style={{ color: MUTED }}>BRING RECORDS IN</div>
          <div className="font-bold text-base mt-0.5 mb-4">Import a spreadsheet</div>
          <div
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
            onClick={() => fileRef.current?.click()}
            className="rounded-2xl border-2 border-dashed flex flex-col items-center justify-center text-center py-12 px-6 cursor-pointer transition-colors"
            style={{ borderColor: dragOver ? "#1F8A4C" : "#D8CFBB", background: dragOver ? "#EDF8F1" : BEIGE }}
          >
            <div className="w-11 h-11 rounded-full bg-white border flex items-center justify-center mb-3" style={{ borderColor: BORDER }}>
              <UploadCloud size={20} style={{ color: dragOver ? "#1F8A4C" : "#1F8A4C" }} />
            </div>
            <div className="text-sm font-semibold">Choose an Excel file (.xlsx)</div>
            <div className="text-xs mt-1" style={{ color: MUTED }}>Click to browse · matches your original workbook layout</div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => handleFiles(e.target.files)} />
          </div>
          <div className="text-xs mt-3 flex items-center gap-1.5" style={{ color: "#C0392B" }}>
            <AlertTriangle size={13} className="shrink-0" />
            Importing replaces all current units, tenants and transactions with what's in the file.
          </div>
        </div>

        <div className="rounded-2xl border bg-white p-5" style={{ borderColor: BORDER }}>
          <div className="text-[11px] font-bold tracking-wide" style={{ color: MUTED }}>TAKE RECORDS OUT</div>
          <div className="font-bold text-base mt-0.5">Build an export</div>
          <p className="text-xs mt-1 mb-4" style={{ color: MUTED }}>Select the records you want to take with you.</p>
          <div className="flex flex-col gap-1 mb-4">
            {items.map(([k, title, desc, count]) => (
              <label key={k} className="flex items-center gap-3 rounded-xl px-2 py-2.5 cursor-pointer hover:bg-black/[0.015]">
                <input type="checkbox" checked={include[k]} onChange={(e) => setInclude((i) => ({ ...i, [k]: e.target.checked }))} className="w-4 h-4 shrink-0" />
                <div className="flex-1 min-w-0">
                  <div className="text-sm font-semibold">{title}</div>
                  <div className="text-[11px]" style={{ color: MUTED }}>{desc}</div>
                </div>
                <div className="text-xs shrink-0" style={{ color: MUTED }}>{count}</div>
              </label>
            ))}
          </div>
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div className="text-xs" style={{ color: MUTED }}>{selectedCount} record{selectedCount === 1 ? "" : "s"} selected</div>
            <div className="flex gap-2 flex-wrap">
              <Btn icon={FileDown} onClick={() => onExportCSV(include)}>Download CSV</Btn>
              <Btn variant="secondary" icon={FileSpreadsheet} onClick={() => onExportExcel(include)}>Download Excel</Btn>
            </div>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <Card>
          <div className="font-bold text-sm mb-1">Start with a template</div>
          <p className="text-xs mb-2" style={{ color: MUTED }}>Use the same column names as your existing workbook for a faster review.</p>
          <button onClick={showColumns} className="text-xs font-semibold" style={{ color: "#1F8A4C" }}>See required columns</button>
        </Card>
        <Card>
          <div className="font-bold text-sm mb-1">Keep your history intact</div>
          <p className="text-xs" style={{ color: MUTED }}>Closed tenancy records never disappear when you migrate your current ledger.</p>
        </Card>
        <Card>
          <div className="font-bold text-sm mb-1">Export when you need to</div>
          <p className="text-xs" style={{ color: MUTED }}>Your data belongs to you. Download a working copy any time.</p>
        </Card>
      </div>

      {pendingFile && (
        <Modal title="Replace all data?" onClose={() => setPendingFile(null)}>
          <p className="text-sm mb-2">
            Importing <span className="font-semibold">{pendingFile.name}</span> will replace all current units, tenants and transactions with what's in this file.
          </p>
          <p className="text-sm mb-5" style={{ color: "#C0392B" }}>
            This can't be undone from within MaliDesk. Consider downloading an export first if you want a backup.
          </p>
          <div className="flex justify-end gap-2">
            <Btn variant="secondary" onClick={() => setPendingFile(null)}>Cancel</Btn>
            <Btn icon={UploadCloud} onClick={confirmImport}>Replace data</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Production authentication shell                                         */
/* ---------------------------------------------------------------------- */
const authApi = async (path, options = {}) => {
  const res = await fetch(path, { credentials: "include", headers: { "Content-Type": "application/json", ...(options.headers || {}) }, ...options });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(body.error || "Request failed");
  return body;
};

function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [remember, setRemember] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const submit = async (e) => {
    e.preventDefault(); setBusy(true); setError("");
    try { const result = await authApi("/api/auth/login", { method: "POST", body: JSON.stringify({ username, password, remember }) }); onLogin(result.user); }
    catch (err) { setError(err.message || "Invalid username or password"); }
    finally { setBusy(false); }
  };
  return <div className="min-h-screen flex items-center justify-center p-5" style={{ background: BEIGE, color: INK }}>
    <div className="w-full max-w-md rounded-3xl border bg-white p-7 shadow-sm" style={{ borderColor: BORDER }}>
      <div className="flex items-center gap-3 mb-7"><div className="w-12 h-12 rounded-2xl flex items-center justify-center font-bold" style={{ background: AMBER, color: SIDEBAR }}>MD</div><div><div className="text-xl font-bold">MaliDesk</div><div className="text-[10px] font-bold tracking-wider" style={{ color: MUTED }}>SECURE RENTAL OPERATIONS</div></div></div>
      <div className="flex items-center gap-2 mb-2"><ShieldCheck size={18} style={{ color: "#1F8A4C" }} /><h1 className="text-lg font-bold">Sign in</h1></div>
      <p className="text-sm mb-5" style={{ color: MUTED }}>Use your MaliDesk account to access the rental portfolio.</p>
      <form onSubmit={submit}>
        <Field label="Username"><input autoComplete="username" value={username} onChange={e => setUsername(e.target.value)} className={inputCls} style={inputStyle} required /></Field>
        <Field label="Password"><div className="relative"><Lock size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: MUTED }} /><input type="password" autoComplete="current-password" value={password} onChange={e => setPassword(e.target.value)} className={`${inputCls} pl-9`} style={inputStyle} required /></div></Field>
        <label className="flex items-center gap-2 text-xs font-semibold mb-4"><input type="checkbox" checked={remember} onChange={e => setRemember(e.target.checked)} /> Remember me on this device</label>
        {error && <div className="text-xs text-red-700 bg-red-50 rounded-xl p-3 mb-4">{error}</div>}
        <button disabled={busy} className="w-full rounded-xl py-2.5 text-sm font-bold disabled:opacity-50" style={{ background: SIDEBAR, color: "white" }}>{busy ? "Signing in…" : "Sign in"}</button>
      </form>
      <div className="text-[11px] mt-5 text-center" style={{ color: MUTED }}>Authentication is handled by the MaliDesk server. Passwords are never stored in this browser.</div>
    </div>
  </div>;
}

function UserManagement({ auth }) {
  const emptyForm = { fullName: "", username: "", email: "", role: "Viewer", status: "active", password: "" };
  const [users, setUsers] = useState([]); const [open, setOpen] = useState(false); const [editing, setEditing] = useState(null); const [busy, setBusy] = useState(false); const [error, setError] = useState(""); const [success, setSuccess] = useState(""); const [form, setForm] = useState(emptyForm);
  const roleDescriptions = [["Administrator", "Full access"],["Manager", "Can manage units, tenants, payments, reports, etc."],["Staff", "Limited operational access"],["Viewer", "Read-only access"]];
  const load = useCallback(async () => { try { const r = await authApi("/api/users"); setUsers(r.users || []); } catch (e) { setError(e.message || "Unable to load users."); } }, []);
  useEffect(() => { load(); }, [load]);
  const openCreate = () => { setEditing(null); setForm(emptyForm); setError(""); setSuccess(""); setOpen(true); };
  const openEdit = (u) => { setEditing(u); setForm({ fullName:u.full_name||"", username:u.username||"", email:u.email||"", role:u.role||"Viewer", status:u.status||"active", password:"" }); setError(""); setSuccess(""); setOpen(true); };
  const save = async (e) => { e.preventDefault(); setBusy(true); setError(""); setSuccess(""); try { if (editing) await authApi(`/api/users/${editing.id}`, {method:"PATCH", body:JSON.stringify(form)}); else await authApi("/api/users", {method:"POST", body:JSON.stringify(form)}); setOpen(false); setEditing(null); setForm(emptyForm); await load(); setSuccess(editing ? "User updated successfully." : "User created successfully."); } catch(e) { setError(e.message || (editing ? "Unable to update user." : "Unable to create user.")); } finally { setBusy(false); } };
  const toggle = async (u) => { setError(""); setSuccess(""); try { await authApi(`/api/users/${u.id}/status`, {method:"PATCH", body:JSON.stringify({status:u.status === "active" ? "inactive" : "active"})}); await load(); setSuccess(`User ${u.status === "active" ? "deactivated" : "activated"} successfully.`); } catch(e) { setError(e.message || "Unable to update status."); } };
  return <div className="pt-2">
    <div className="flex items-start justify-between gap-3 mb-5"><div><div className="text-[11px] font-bold tracking-wide" style={{color:MUTED}}>SECURITY</div><h1 className="text-xl font-bold">Users & access</h1><p className="text-sm" style={{color:MUTED}}>Manage accounts, roles and access status.</p></div><Btn icon={UserPlus} onClick={openCreate}>Add User</Btn></div>
    <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-3 mb-5">{roleDescriptions.map(([name,description])=><div key={name} className="rounded-xl border bg-white p-3" style={{borderColor:BORDER}}><div className="font-bold text-sm">{name}</div><div className="text-[11px] mt-1" style={{color:MUTED}}>{description}</div></div>)}</div>
    {error && <div className="text-sm text-red-700 bg-red-50 rounded-xl p-3 mb-4">{error}</div>}{success && <div className="text-sm text-green-700 bg-green-50 rounded-xl p-3 mb-4">{success}</div>}
    <div className="rounded-2xl border bg-white overflow-hidden" style={{borderColor:BORDER}}><div className="overflow-x-auto"><table className="w-full text-sm"><thead><tr style={{borderBottom:`1px solid ${BORDER}`}}>{["Full Name","Username","Email","Role","Status","Last Login","Date Created","Actions"].map(h=><th key={h} className="text-left px-3 py-2.5 text-xs font-bold whitespace-nowrap" style={{color:MUTED}}>{h}</th>)}</tr></thead><tbody>{users.map(u=><tr key={u.id} style={{borderBottom:`1px solid ${BORDER}`}}><td className="px-3 py-3 font-semibold">{u.full_name}</td><td className="px-3 py-3">{u.username}</td><td className="px-3 py-3">{u.email||"—"}</td><td className="px-3 py-3 font-semibold">{u.role}</td><td className="px-3 py-3"><Pill label={u.status === "active" ? "Active" : "Inactive"} colorKey={u.status === "active" ? "PAID" : "HIGH"}/></td><td className="px-3 py-3">{u.last_login_at ? fmtDate(u.last_login_at) : "Never"}</td><td className="px-3 py-3">{fmtDate(u.created_at)}</td><td className="px-3 py-3"><div className="flex items-center gap-3"><button onClick={()=>openEdit(u)} className="text-xs font-semibold" style={{color:AMBER_DARK}}><Pencil size={13} className="inline mr-1"/>Edit</button><button onClick={()=>toggle(u)} className="text-xs font-semibold" style={{color:u.status === "active" ? "#C0392B" : "#1F8A4C"}}>{u.status === "active" ? "Deactivate" : "Activate"}</button></div></td></tr>)}</tbody></table></div></div>
    {open && <Modal title={editing ? "Edit User" : "Add User"} onClose={()=>{if(!busy){setOpen(false);setEditing(null)}}}><form onSubmit={save}>
      <Field label="Full Name"><input value={form.fullName} onChange={e=>setForm({...form,fullName:e.target.value})} className={inputCls} style={inputStyle} required/></Field><Field label="Username"><input autoComplete="username" value={form.username} onChange={e=>setForm({...form,username:e.target.value})} className={inputCls} style={inputStyle} required/></Field><Field label="Email"><input type="email" value={form.email} onChange={e=>setForm({...form,email:e.target.value})} className={inputCls} style={inputStyle}/></Field><Field label="Role"><select value={form.role} onChange={e=>setForm({...form,role:e.target.value})} className={inputCls} style={inputStyle}><option value="Administrator">Administrator — Full access</option><option value="Manager">Manager — Units, tenants, payments, reports, etc.</option><option value="Staff">Staff — Limited operational access</option><option value="Viewer">Viewer — Read-only access</option></select></Field>{editing&&<Field label="Status"><select value={form.status} onChange={e=>setForm({...form,status:e.target.value})} className={inputCls} style={inputStyle}><option value="active">Active</option><option value="inactive">Inactive</option></select></Field>}<Field label={editing?"New password (optional)":"Temporary password"} hint={editing?"Leave blank to keep the current password.":"At least 8 characters with upper, lower and number."}><input type="password" autoComplete="new-password" value={form.password} onChange={e=>setForm({...form,password:e.target.value})} className={inputCls} style={inputStyle} required={!editing}/></Field>
      {error&&<div className="text-sm text-red-700 bg-red-50 rounded-xl p-3 mb-4">{error}</div>}<div className="flex justify-end gap-2"><Btn variant="secondary" type="button" onClick={()=>setOpen(false)}>Cancel</Btn><Btn type="submit" disabled={busy}>{busy?(editing?"Saving…":"Creating…"):(editing?"Save Changes":"Create User")}</Btn></div></form></Modal>}
  </div>;
}

export default function MaliDesk() {
  const [user, setUser] = useState(null);
  const [checking, setChecking] = useState(true);
  useEffect(() => { authApi("/api/auth/me").then(r => setUser(r.user)).catch(() => {}).finally(() => setChecking(false)); }, []);
  const logout = async () => { try { await authApi("/api/auth/logout", {method:"POST"}); } finally { setUser(null); } };
  if (checking) return <div className="min-h-screen flex items-center justify-center" style={{background:BEIGE,color:MUTED}}>Checking secure session…</div>;
  if (!user) return <LoginScreen onLogin={setUser} />;
  return <MaliDeskCore auth={{user,logout}} />;
}
