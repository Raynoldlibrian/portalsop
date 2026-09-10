import { useState, useMemo, useEffect, useRef } from "react";
import {
  Search,
  FileText,
  Download,
  Building2,
  CalendarCheck,
  ChevronDown,
  X,
  Sparkles,
  ClipboardList,
  Landmark,
  Layers,
  Plus,
  Link2,
  Send,
  CheckCircle2,
  MoreVertical,
  Pencil,
  History,
  ArrowLeft,
  Eye,
  ShieldCheck,
  ShieldAlert,
  BookOpen,
  FileSpreadsheet,
  Printer,
} from "lucide-react";
import * as XLSX from "xlsx";

// ==================== KONFIGURASI API ====================
const API_BASE_URL =
  "https://script.google.com/macros/s/AKfycbyiStxNAee_or-ABISDOZY-zsh3ZtGcRh4OiUQTSVkHiQQNaEiuBggKuJj9YrCnJEO9/exec";

async function apiGet(action, params) {
  const qs = new URLSearchParams({ action, ...(params || {}) }).toString();
  const res = await fetch(`${API_BASE_URL}?${qs}`);
  const json = await res.json();
  if (!json.success) throw new Error(json.message || "Gagal mengambil data");
  return json.data;
}

async function apiPost(action, payload) {
  const res = await fetch(API_BASE_URL, {
    method: "POST",
    // text/plain menghindari CORS preflight yang tidak didukung Apps Script Web App
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action, ...payload }),
  });
  const json = await res.json();
  if (!json.success) throw new Error(json.message || "Gagal mengirim data");
  return json.data;
}

// Ubah bentuk objek SOP dari API (snake_case, sesuai kolom Sheet) ke bentuk
// yang dipakai komponen di file ini (camelCase, sudah dipakai sejak versi mockup awal)
function mapSopFromApi(s) {
  const riwayat = (s.riwayat_revisi || []).map((r) => ({
    tanggal: r.tanggal,
    keterangan: r.keterangan,
  }));
  const tglRevisiTerakhir = riwayat.length > 0 ? riwayat[riwayat.length - 1].tanggal : "—";
  return {
    id: s.id_sop,
    idOpd: s.id_opd,
    nomor: s.nomor_sop,
    judul: s.judul_sop,
    opd: s.nama_opd,
    bidang: s.bidang_bagian,
    seksi: s.seksi_subbid_subbag,
    abstrak: s.abstrak,
    tglPembuatan: s.tgl_pembuatan,
    tglRevisi: tglRevisiTerakhir,
    tglEfektif: s.tgl_efektif,
    disahkanOleh: s.disahkan_oleh,
    status: s.status,
    linkDrive: s.link_drive,
    statusVerifikasi: s.status_verifikasi,
    catatanVerifikasi: s.catatan_verifikasi,
    riwayat,
    diverifikasi: s.tgl_verifikasi || s.tgl_submit || "",
  };
}


function Chip({ children, onRemove }) {
  return (
    <span className="inline-flex items-center gap-1.5 rounded-full bg-teal-800 text-teal-50 text-xs font-medium pl-3 pr-2 py-1">
      {children}
      <button
        onClick={onRemove}
        className="rounded-full hover:bg-white/15 p-0.5 transition-colors"
        aria-label={`Hapus filter ${children}`}
      >
        <X size={11} strokeWidth={2.5} />
      </button>
    </span>
  );
}

function CardMenu({ onEdit, onRevisi }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="relative shrink-0">
      <button
        onClick={() => setOpen((v) => !v)}
        className="w-8 h-8 rounded flex items-center justify-center text-gray-500 hover:bg-teal-50 transition-colors"
        aria-label="Menu SOP"
      >
        <MoreVertical size={16} strokeWidth={2.2} />
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-10" onClick={() => setOpen(false)} />
          <div className="absolute right-0 z-20 mt-1 w-40 bg-white border border-stone-200 rounded-md shadow-lg py-1">
            <button
              onClick={() => {
                setOpen(false);
                onEdit();
              }}
              className="w-full flex items-center gap-2 text-left px-3 py-2 text-xs text-gray-700 hover:bg-stone-50"
            >
              <Pencil size={13} strokeWidth={2} />
              Edit SOP
            </button>
            <button
              onClick={() => {
                setOpen(false);
                onRevisi();
              }}
              className="w-full flex items-center gap-2 text-left px-3 py-2 text-xs text-gray-700 hover:bg-stone-50"
            >
              <History size={13} strokeWidth={2} />
              Catat Revisi
            </button>
          </div>
        </>
      )}
    </div>
  );
}

function SopCard({ sop, onEdit, onRevisi, onOpen, canManage }) {
  return (
    <article
      onClick={onOpen}
      className="group relative bg-white border border-stone-200 rounded-xl overflow-hidden transition-shadow hover:shadow-lg cursor-pointer"
    >
      <div className="px-5 py-5">
        <div className="flex items-start justify-between gap-4">
          <div className="min-w-0">
            <div className="flex items-center gap-2 text-xs tracking-wider uppercase text-teal-800 font-semibold mb-1.5">
              <span className="font-mono tracking-normal text-xs bg-teal-50 border border-stone-200 rounded px-1.5 py-0.5 text-teal-900">
                {sop.nomor}
              </span>
              <span className="text-stone-300">·</span>
              <span className="normal-case font-medium text-gray-500 flex items-center gap-1">
                <Building2 size={12} strokeWidth={2} />
                {sop.opd}
              </span>
            </div>
            <h3 className="text-lg leading-snug font-semibold text-teal-900 mb-2" style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}>
              {sop.judul}
            </h3>
            <BadgePerluDitinjau tglEfektif={sop.tglEfektif} />
            <dl className="grid grid-cols-3 gap-x-4 gap-y-1 text-xs mt-2">
              <div>
                <dt className="text-stone-400">Efektif</dt>
                <dd className="text-gray-700 font-medium">{formatTanggal(sop.tglEfektif)}</dd>
              </div>
              <div>
                <dt className="text-stone-400">Revisi</dt>
                <dd className="text-gray-700 font-medium">{formatTanggal(sop.tglRevisi)}</dd>
              </div>
              <div className="col-span-1">
                <dt className="text-stone-400">Disahkan oleh</dt>
                <dd className="text-gray-700 font-medium truncate">{sop.disahkanOleh}</dd>
              </div>
            </dl>
          </div>
          <div className="shrink-0 flex items-start gap-1" onClick={(e) => e.stopPropagation()}>
            {sop.linkDrive ? (
              <a
                href={sop.linkDrive}
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-1.5 rounded border border-teal-800 text-teal-900 text-xs font-semibold px-3 py-2 hover:bg-teal-700 hover:text-white transition-colors"
              >
                <Download size={13} strokeWidth={2.3} />
                Lihat
              </a>
            ) : (
              <span className="inline-flex items-center gap-1.5 rounded border border-stone-200 text-stone-300 text-xs font-semibold px-3 py-2 cursor-not-allowed">
                <Download size={13} strokeWidth={2.3} />
                Lihat
              </span>
            )}
            {canManage && <CardMenu onEdit={onEdit} onRevisi={onRevisi} />}
          </div>
        </div>
      </div>
    </article>
  );
}

function MetaRow({ label, children, index }) {
  return (
    <div className={`grid grid-cols-3 gap-4 px-5 py-3 ${index % 2 === 1 ? "bg-stone-50" : "bg-white"}`}>
      <dt className="text-sm font-semibold text-gray-700">{label}</dt>
      <dd className="col-span-2 text-sm text-gray-700">{children}</dd>
    </div>
  );
}

function SopDetailPage({ sop, isOperator, onBack, onEdit, onRevisi, verifikasiPanel, backLabel = "Kembali ke pencarian" }) {
  const berlaku = sop.status === "Berlaku";
  const [catatanVerifikasi, setCatatanVerifikasi] = useState("");
  const [verifBusy, setVerifBusy] = useState(false);
  const [verifError, setVerifError] = useState("");

  const kirimVerifikasi = async (statusVerifikasi) => {
    if (statusVerifikasi === "Revisi" && !catatanVerifikasi.trim()) {
      setVerifError("Isi catatan dulu, biar OPD tau apa yang perlu diperbaiki");
      return;
    }
    setVerifError("");
    setVerifBusy(true);
    try {
      await apiPost("verifikasi", {
        id_sop: sop.id,
        status_verifikasi: statusVerifikasi,
        catatan_verifikasi: catatanVerifikasi.trim(),
        verifikator: verifikasiPanel?.verifikatorName,
      });
      verifikasiPanel?.onSelesai();
    } catch (err) {
      setVerifError(err.message || "Gagal mengirim, coba lagi.");
      setVerifBusy(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto px-6 py-8">
      <button
        onClick={onBack}
        className="inline-flex items-center gap-1.5 text-sm font-medium text-gray-500 hover:text-teal-800 transition-colors mb-5"
      >
        <ArrowLeft size={15} strokeWidth={2.2} />
        {backLabel}
      </button>

      {isOperator && (
        <div className="mb-5 flex items-center justify-between gap-3 bg-amber-50 border border-amber-200 rounded-xl px-4 py-3">
          <span className="text-xs text-gray-700">
            Anda masuk sebagai operator <strong>{sop.opd}</strong>.
          </span>
          <div className="flex items-center gap-2 shrink-0">
            <button
              onClick={onEdit}
              className="inline-flex items-center gap-1.5 text-xs font-semibold text-teal-900 bg-white border border-stone-200 rounded-md px-3 py-1.5 hover:border-teal-800 transition-colors"
            >
              <Pencil size={12} strokeWidth={2.2} />
              Edit SOP
            </button>
            <button
              onClick={onRevisi}
              className="inline-flex items-center gap-1.5 text-xs font-semibold text-white bg-teal-800 rounded-md px-3 py-1.5 hover:bg-teal-700 transition-colors"
            >
              <History size={12} strokeWidth={2.2} />
              Catat Revisi
            </button>
          </div>
        </div>
      )}

      {/* Title block */}
      <div className="bg-teal-900 rounded-xl px-6 py-6 mb-5">
        <div className="flex items-center gap-2 text-xs font-semibold tracking-widest uppercase text-teal-200 mb-2">
          <FileText size={13} strokeWidth={2.3} />
          Standar Operasional Prosedur
        </div>
        <h1
          className="text-2xl leading-snug font-bold text-white mb-3"
          style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
        >
          {sop.judul}
        </h1>
        <div className="flex flex-wrap items-center gap-2">
          <span className="font-mono text-xs bg-white bg-opacity-10 text-teal-100 rounded px-2 py-1">
            {sop.nomor}
          </span>
          <span
            className={`inline-flex items-center gap-1 text-xs font-semibold rounded-full px-2.5 py-1 ${
              berlaku ? "bg-emerald-100 text-emerald-800" : "bg-red-100 text-red-700"
            }`}
          >
            {berlaku ? <ShieldCheck size={12} strokeWidth={2.3} /> : <ShieldAlert size={12} strokeWidth={2.3} />}
            {sop.status}
          </span>
          {perluDitinjau(sop.tglEfektif) && (
            <span className="inline-flex items-center gap-1 text-xs font-semibold rounded-full px-2.5 py-1 bg-amber-100 text-amber-800">
              <History size={12} strokeWidth={2.3} />
              Perlu ditinjau ulang ({Math.floor(tahunSejakEfektif(sop.tglEfektif))} tahun sejak efektif)
            </span>
          )}
        </div>
      </div>

      {sop.abstrak && (
        <div className="bg-white border border-stone-200 rounded-xl px-5 py-5 mb-5">
          <div className="flex items-center gap-2 text-sm font-semibold text-teal-900 mb-2">
            <BookOpen size={15} strokeWidth={2.2} />
            Abstraksi
          </div>
          <p className="text-sm text-gray-700 leading-relaxed">{sop.abstrak}</p>
        </div>
      )}

      {/* File action card */}
      <div className="bg-white border border-stone-200 rounded-xl px-5 py-5 mb-5">
        <div className="flex items-center gap-2 text-sm font-semibold text-teal-900 mb-3">
          <FileText size={15} strokeWidth={2.2} />
          Dokumen SOP
        </div>
        <div className="flex flex-wrap items-center gap-2">
          {sop.linkDrive ? (
            <>
              <a
                href={sop.linkDrive}
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-1.5 text-sm font-semibold text-white bg-teal-800 rounded-md px-4 py-2.5 hover:bg-teal-700 transition-colors"
              >
                <Eye size={14} strokeWidth={2.2} />
                Lihat PDF
              </a>
              <a
                href={sop.linkDrive}
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-1.5 text-sm font-semibold text-teal-900 border border-teal-800 rounded-md px-4 py-2.5 hover:bg-teal-50 transition-colors"
              >
                <Download size={14} strokeWidth={2.2} />
                Unduh
              </a>
            </>
          ) : (
            <p className="text-sm text-stone-400">Link dokumen belum tersedia.</p>
          )}
        </div>
      </div>

      {/* Metadata */}
      <div className="bg-white border border-stone-200 rounded-xl overflow-hidden mb-5">
        <div className="px-5 py-3 border-b border-stone-200">
          <h2 className="text-sm font-semibold text-teal-900">Metadata SOP</h2>
        </div>
        <MetaRow label="Judul" index={0}>
          {sop.judul}
        </MetaRow>
        <MetaRow label="Nomor SOP" index={1}>
          <span className="font-mono">{sop.nomor}</span>
        </MetaRow>
        <MetaRow label="OPD" index={2}>
          {sop.opd}
        </MetaRow>
        <MetaRow label="Bidang / Bagian" index={3}>
          {sop.bidang}
        </MetaRow>
        <MetaRow label="Seksi / Subbid / Subbag" index={4}>
          {sop.seksi}
        </MetaRow>
        <MetaRow label="Tanggal Pembuatan" index={5}>
          {formatTanggal(sop.tglPembuatan)}
        </MetaRow>
        <MetaRow label="Tanggal Revisi" index={6}>
          {formatTanggal(sop.tglRevisi)}
          {sop.riwayat.length > 1 && (
            <ul className="mt-2 space-y-1">
              {sop.riwayat.map((r, i) => (
                <li key={i} className="flex items-start gap-2 text-xs text-gray-500">
                  <History size={11} strokeWidth={2} className="mt-0.5 shrink-0" />
                  <span>
                    {formatTanggal(r.tanggal)} — {r.keterangan}
                  </span>
                </li>
              ))}
            </ul>
          )}
        </MetaRow>
        <MetaRow label="Tanggal Efektif" index={7}>
          {formatTanggal(sop.tglEfektif)}
        </MetaRow>
        <MetaRow label="Pejabat yang Mengesahkan" index={8}>
          {sop.disahkanOleh}
        </MetaRow>
        <MetaRow label="Status" index={9}>
          <span className={berlaku ? "text-emerald-700 font-semibold" : "text-red-600 font-semibold"}>
            {sop.status}
          </span>
        </MetaRow>
      </div>

      {verifikasiPanel && (
        <div className="bg-white border border-stone-200 rounded-xl px-5 py-5 mb-5">
          <div className="flex items-center gap-2 text-sm font-semibold text-teal-900 mb-3">
            <ShieldCheck size={15} strokeWidth={2.2} />
            Verifikasi
          </div>
          <label className="block text-xs font-semibold text-gray-700 mb-1.5">
            Catatan Revisi <span className="text-stone-400 font-normal">(wajib kalau minta revisi)</span>
          </label>
          <textarea
            rows={3}
            value={catatanVerifikasi}
            onChange={(e) => setCatatanVerifikasi(e.target.value)}
            placeholder="cth. Link Drive belum bisa diakses publik…"
            className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800 resize-none"
          />
          {verifError && <p className="mt-2 text-xs text-red-600">{verifError}</p>}
          <div className="mt-3 flex items-center justify-end gap-2">
            <button
              type="button"
              disabled={verifBusy}
              onClick={() => kirimVerifikasi("Revisi")}
              className="inline-flex items-center gap-1.5 text-sm font-semibold text-amber-700 border border-amber-400 rounded-md px-4 py-2 hover:bg-amber-50 transition-colors disabled:opacity-50"
            >
              <History size={14} strokeWidth={2.2} />
              Revisi
            </button>
            <button
              type="button"
              disabled={verifBusy}
              onClick={() => kirimVerifikasi("Terverifikasi")}
              className="inline-flex items-center gap-1.5 text-sm font-semibold text-white bg-teal-800 rounded-md px-4 py-2 hover:bg-teal-700 transition-colors disabled:opacity-50"
            >
              <CheckCircle2 size={14} strokeWidth={2.2} />
              Verifikasi
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

const STAT_COLORS = {
  emerald: { bg: "bg-emerald-50", icon: "text-emerald-700", value: "text-emerald-800" },
  blue: { bg: "bg-sky-50", icon: "text-sky-700", value: "text-sky-800" },
  slate: { bg: "bg-stone-100", icon: "text-stone-600", value: "text-stone-800" },
  amber: { bg: "bg-amber-50", icon: "text-amber-700", value: "text-amber-800" },
};

function StatCard({ icon: Icon, label, value, color, onClick }) {
  const c = STAT_COLORS[color] || STAT_COLORS.blue;
  const Tag = onClick ? "button" : "div";
  return (
    <Tag
      onClick={onClick}
      className={`flex flex-col items-start gap-2 bg-white border border-stone-200 rounded-xl px-3 py-3 sm:px-4 sm:py-4 min-w-0 text-left w-full ${
        onClick ? "hover:border-teal-800 hover:shadow-md transition-all cursor-pointer" : ""
      }`}
    >
      <div className={`w-8 h-8 rounded-md ${c.bg} flex items-center justify-center`}>
        <Icon size={15} strokeWidth={2} className={c.icon} />
      </div>
      <div className="min-w-0">
        <div
          className={`text-xl sm:text-2xl leading-none font-semibold ${c.value} mb-1 truncate`}
          style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
        >
          {value}
        </div>
        <div className="text-xs sm:text-xs leading-snug text-gray-500">{label}</div>
      </div>
    </Tag>
  );
}

function OpdListPage({ opdCounts, onBack, onSelectOpd }) {
  return (
    <div className="max-w-4xl mx-auto px-6 py-8">
      <button
        onClick={onBack}
        className="inline-flex items-center gap-1.5 text-sm font-medium text-gray-500 hover:text-teal-800 transition-colors mb-5"
      >
        <ArrowLeft size={15} strokeWidth={2.2} />
        Kembali ke pencarian
      </button>

      <div className="bg-teal-900 rounded-xl px-6 py-6 mb-5">
        <div className="text-xs font-semibold tracking-widest uppercase text-teal-200 mb-2">
          Registri SOP
        </div>
        <h1
          className="text-2xl leading-snug font-bold text-white"
          style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
        >
          OPD Berpartisipasi
        </h1>
        <p className="text-sm text-teal-100 mt-2">
          {opdCounts.length} OPD sudah punya SOP terverifikasi di portal ini
        </p>
      </div>

      {opdCounts.length === 0 ? (
        <div className="text-center py-20 border border-dashed border-stone-200 rounded-md">
          <Landmark size={28} strokeWidth={1.5} className="mx-auto text-teal-100 mb-3" />
          <p className="text-gray-500 text-sm">Belum ada OPD dengan SOP terverifikasi.</p>
        </div>
      ) : (
        <div className="grid gap-2">
          {opdCounts.map((o) => (
            <button
              key={o.opd}
              onClick={() => onSelectOpd(o.opd)}
              className="w-full flex items-center justify-between gap-3 bg-white border border-stone-200 rounded-xl px-5 py-4 hover:border-teal-800 hover:shadow-md transition-all text-left"
            >
              <div className="flex items-center gap-3 min-w-0">
                <div className="w-9 h-9 rounded-md bg-sky-50 flex items-center justify-center shrink-0">
                  <Landmark size={16} strokeWidth={2} className="text-sky-700" />
                </div>
                <span className="text-sm font-semibold text-teal-900 truncate">{o.opd}</span>
              </div>
              <span className="shrink-0 inline-flex items-center gap-1.5 text-xs font-semibold text-emerald-800 bg-emerald-50 rounded-full px-2.5 py-1">
                {o.count} SOP
              </span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function LatestSopCard({ sop, onOpen }) {
  return (
    <div
      onClick={onOpen}
      className="relative shrink-0 w-64 bg-white border border-stone-200 rounded-xl p-4 snap-start cursor-pointer hover:shadow-lg transition-shadow"
    >
      <div className="absolute left-0 top-3 bottom-3 w-1 bg-amber-500 rounded-r" />
      <div className="pl-3">
        <span className="font-mono text-xs bg-teal-50 border border-stone-200 rounded px-1.5 py-0.5 text-teal-900">
          {sop.nomor}
        </span>
        <h4
          className="mt-2 text-sm leading-snug font-semibold text-teal-900 line-clamp-2"
          style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
        >
          {sop.judul}
        </h4>
        <div className="mt-2 flex items-center gap-1.5 text-xs text-gray-500">
          <Building2 size={11} strokeWidth={2} />
          <span className="truncate">{sop.opd}</span>
        </div>
      </div>
    </div>
  );
}

// mode: "tambah" | "edit" | "revisi". initialSop is null for "tambah".
function LoginModal({ opdList, onClose, onLogin }) {
  const [idOpd, setIdOpd] = useState("");
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async () => {
    if (!idOpd) {
      setError("Pilih OPD dulu");
      return;
    }
    if (pin.trim().length !== 4) {
      setError("PIN harus 4 digit");
      return;
    }
    setError("");
    setLoading(true);
    try {
      const data = await apiPost("login", { id_opd: idOpd, pin });
      onLogin(data); // { id_opd, nama_opd, singkatan }
    } catch (err) {
      setError(err.message || "PIN salah");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center px-4">
      <div className="absolute inset-0 bg-teal-800 bg-opacity-40 backdrop-blur-sm" onClick={onClose} />
      <div className="relative w-full max-w-sm bg-stone-50 rounded-2xl shadow-2xl overflow-hidden">
        <div className="border-b border-stone-200 px-6 py-4 flex items-center justify-between">
          <div>
            <div className="text-xs font-semibold tracking-widest uppercase text-teal-700">Login Operator</div>
            <h2
              className="text-xl font-semibold text-teal-900"
              style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
            >
              Masuk sebagai OPD
            </h2>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full flex items-center justify-center text-gray-500 hover:bg-teal-50 transition-colors"
            aria-label="Tutup"
          >
            <X size={16} strokeWidth={2.2} />
          </button>
        </div>

        <div className="px-6 py-5 space-y-4">
          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">Pilih OPD</label>
            <select
              value={idOpd}
              onChange={(e) => {
                setIdOpd(e.target.value);
                setError("");
              }}
              className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            >
              <option value="">Pilih OPD…</option>
              {opdList.map((o) => (
                <option key={o.id_opd} value={o.id_opd}>
                  {o.nama_opd}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">PIN Akses OPD</label>
            <input
              type="password"
              inputMode="numeric"
              maxLength={4}
              value={pin}
              onChange={(e) => setPin(e.target.value.replace(/[^0-9]/g, ""))}
              placeholder="••••"
              className="w-28 rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm tracking-widest text-center text-teal-900 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
            <p className="mt-1.5 text-xs text-gray-500">Minta PIN akses OPD ke Bagian Organisasi jika belum punya.</p>
          </div>

          {error && <p className="text-xs text-red-600">{error}</p>}

          <div className="pt-2 flex items-center justify-end gap-2">
            <button
              type="button"
              onClick={onClose}
              className="text-sm font-medium text-gray-500 px-4 py-2.5 hover:text-gray-700 transition-colors"
            >
              Batal
            </button>
            <button
              type="button"
              onClick={handleSubmit}
              disabled={loading}
              className="inline-flex items-center gap-2 text-sm font-semibold text-white bg-teal-800 rounded-md px-4 py-2.5 hover:bg-teal-700 transition-colors disabled:opacity-50"
            >
              <ShieldCheck size={14} strokeWidth={2.2} />
              {loading ? "Memeriksa…" : "Masuk"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

const BULAN_ID = {
  januari: "01", februari: "02", maret: "03", april: "04", mei: "05", juni: "06",
  juli: "07", agustus: "08", september: "09", oktober: "10", november: "11", desember: "12",
};

// "3 Agustus 2015" -> "2015-08-03" (untuk defaultValue input type=date). "—" atau kosong -> "".
// Terima format "yyyy-MM-dd" (dari API, sudah pas buat <input type=date>) atau
// format Indonesia lama "3 Agustus 2015" (jaga-jaga data lama). "—"/kosong -> "".
function toIsoDate(tanggal) {
  if (!tanggal || tanggal === "—") return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) return tanggal;
  const parts = tanggal.trim().toLowerCase().split(/\s+/);
  if (parts.length !== 3) return "";
  const [day, bulan, year] = parts;
  const month = BULAN_ID[bulan];
  if (!month) return "";
  return `${year}-${month}-${day.padStart(2, "0")}`;
}

const BULAN_NAMA = [
  "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  "Juli", "Agustus", "September", "Oktober", "November", "Desember",
];

// Kebalikan dari toIsoDate: "2020-02-14" -> "14 Februari 2020", buat tampilan teks.
// Kalau formatnya nggak dikenali (mis. sudah "—" atau string bebas), tampilkan apa adanya.
function formatTanggal(tanggal) {
  if (!tanggal) return "—";
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(tanggal);
  if (!m) return tanggal;
  const [, year, month, day] = m;
  const nama = BULAN_NAMA[parseInt(month, 10) - 1];
  if (!nama) return tanggal;
  return `${parseInt(day, 10)} ${nama} ${year}`;
}

// Unduh daftar SOP yang lagi tampil (sudah kefilter) sebagai file Excel.
function exportExcel(sopArray) {
  const rows = sopArray.map((s) => ({
    "Nomor SOP": s.nomor,
    "Judul SOP": s.judul,
    OPD: s.opd,
    "Bidang/Bagian": s.bidang,
    "Seksi/Subbid/Subbag": s.seksi,
    "Tanggal Pembuatan": formatTanggal(s.tglPembuatan),
    "Tanggal Revisi": formatTanggal(s.tglRevisi),
    "Tanggal Efektif": formatTanggal(s.tglEfektif),
    "Disahkan Oleh": s.disahkanOleh,
    Status: s.status,
    "Perlu Ditinjau": perluDitinjau(s.tglEfektif) ? "Ya" : "Tidak",
    "Link Drive": s.linkDrive,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [
    { wch: 14 }, { wch: 40 }, { wch: 22 }, { wch: 22 }, { wch: 22 },
    { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 22 }, { wch: 12 }, { wch: 12 }, { wch: 40 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Rekap SOP");
  const tanggal = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Rekap-SOP-Inhu-${tanggal}.xlsx`);
}

// SOP dianggap perlu ditinjau ulang kalau tgl_efektif sudah lewat 2 tahun
// (standar peninjauan ulang SOP AP sesuai Permenpan RB).
function tahunSejakEfektif(tglEfektif) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(tglEfektif || "");
  if (!m) return null;
  const efektif = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const now = new Date();
  return (now - efektif) / (1000 * 60 * 60 * 24 * 365.25);
}
function perluDitinjau(tglEfektif) {
  const tahun = tahunSejakEfektif(tglEfektif);
  return tahun !== null && tahun >= 2;
}

function BadgePerluDitinjau({ tglEfektif }) {
  if (!perluDitinjau(tglEfektif)) return null;
  const tahun = tahunSejakEfektif(tglEfektif);
  return (
    <span className="inline-flex items-center gap-1 text-xs font-semibold rounded-full px-2.5 py-1 border bg-amber-50 text-amber-700 border-amber-200">
      <History size={12} strokeWidth={2.3} />
      Perlu ditinjau ({Math.floor(tahun)} thn)
    </span>
  );
}

function VerifikatorLoginModal({ onClose, onLogin }) {
  const [nama, setNama] = useState("");
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async () => {
    if (!nama.trim()) {
      setError("Isi nama dulu");
      return;
    }
    if (pin.trim().length !== 4) {
      setError("PIN harus 4 digit");
      return;
    }
    setError("");
    setLoading(true);
    try {
      const data = await apiPost("verifikatorLogin", { nama: nama.trim(), pin });
      onLogin(data.nama);
    } catch (err) {
      setError(err.message || "Nama atau PIN salah");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center px-4">
      <div className="absolute inset-0 bg-teal-800 bg-opacity-40 backdrop-blur-sm" onClick={onClose} />
      <div className="relative w-full max-w-sm bg-stone-50 rounded-2xl shadow-2xl overflow-hidden">
        <div className="border-b border-stone-200 px-6 py-4 flex items-center justify-between">
          <div>
            <div className="text-xs font-semibold tracking-widest uppercase text-teal-700">
              Login Verifikator
            </div>
            <h2
              className="text-xl font-semibold text-teal-900"
              style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
            >
              Bagian Organisasi
            </h2>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full flex items-center justify-center text-gray-500 hover:bg-teal-50 transition-colors"
            aria-label="Tutup"
          >
            <X size={16} strokeWidth={2.2} />
          </button>
        </div>

        <div className="px-6 py-5 space-y-4">
          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">Nama</label>
            <input
              type="text"
              value={nama}
              onChange={(e) => {
                setNama(e.target.value);
                setError("");
              }}
              placeholder="Nama verifikator"
              className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">PIN</label>
            <input
              type="password"
              inputMode="numeric"
              maxLength={4}
              value={pin}
              onChange={(e) => setPin(e.target.value.replace(/[^0-9]/g, ""))}
              placeholder="••••"
              className="w-28 rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm tracking-widest text-center text-teal-900 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
          </div>

          {error && <p className="text-xs text-red-600">{error}</p>}

          <div className="pt-2 flex items-center justify-end gap-2">
            <button
              type="button"
              onClick={onClose}
              className="text-sm font-medium text-gray-500 px-4 py-2.5 hover:text-gray-700 transition-colors"
            >
              Batal
            </button>
            <button
              type="button"
              onClick={handleSubmit}
              disabled={loading}
              className="inline-flex items-center gap-2 text-sm font-semibold text-white bg-teal-800 rounded-md px-4 py-2.5 hover:bg-teal-700 transition-colors disabled:opacity-50"
            >
              <ShieldCheck size={14} strokeWidth={2.2} />
              {loading ? "Memeriksa…" : "Masuk"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function PendingSopRow({ sop, onOpen }) {
  return (
    <button
      onClick={onOpen}
      className="w-full text-left bg-white border border-stone-200 rounded-xl px-5 py-4 hover:shadow-md hover:border-teal-800 transition-all flex items-center justify-between gap-4"
    >
      <div className="min-w-0">
        <div className="flex items-center gap-2 text-xs tracking-wider uppercase text-teal-800 font-semibold mb-1.5">
          <span className="font-mono tracking-normal text-xs bg-teal-50 border border-stone-200 rounded px-1.5 py-0.5 text-teal-900">
            {sop.nomor}
          </span>
          <span className="text-stone-300">·</span>
          <span className="normal-case font-medium text-gray-500 flex items-center gap-1">
            <Building2 size={12} strokeWidth={2} />
            {sop.opd}
          </span>
        </div>
        <h3
          className="text-sm leading-snug font-semibold text-teal-900 truncate"
          style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
        >
          {sop.judul}
        </h3>
        <p className="text-xs text-stone-400 mt-1">Efektif {formatTanggal(sop.tglEfektif)} · Diajukan oleh {sop.disahkanOleh}</p>
      </div>
      <ChevronDown size={16} strokeWidth={2.3} className="-rotate-90 text-stone-300 shrink-0" />
    </button>
  );
}

const STATUS_BADGE = {
  Menunggu: { bg: "bg-amber-50", text: "text-amber-700", border: "border-amber-200", icon: History },
  Terverifikasi: { bg: "bg-emerald-50", text: "text-emerald-700", border: "border-emerald-200", icon: CheckCircle2 },
  Revisi: { bg: "bg-red-50", text: "text-red-700", border: "border-red-200", icon: ShieldAlert },
};

function StatusBadge({ status }) {
  const s = STATUS_BADGE[status] || STATUS_BADGE.Menunggu;
  const Icon = s.icon;
  return (
    <span
      className={`inline-flex items-center gap-1 text-xs font-semibold rounded-full px-2.5 py-1 border ${s.bg} ${s.text} ${s.border}`}
    >
      <Icon size={12} strokeWidth={2.3} />
      {status}
    </span>
  );
}

function SopSayaCard({ sop, onEdit, onRevisi }) {
  return (
    <div className="bg-white border border-stone-200 rounded-xl p-5">
      <div className="flex items-start justify-between gap-3 mb-2">
        <div className="flex items-center gap-2 text-xs tracking-wider uppercase text-teal-800 font-semibold">
          <span className="font-mono tracking-normal text-xs bg-teal-50 border border-stone-200 rounded px-1.5 py-0.5 text-teal-900">
            {sop.nomor}
          </span>
        </div>
        <StatusBadge status={sop.statusVerifikasi} />
      </div>
      <h3
        className="text-base leading-snug font-semibold text-teal-900 mb-2"
        style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
      >
        {sop.judul}
      </h3>
      <BadgePerluDitinjau tglEfektif={sop.tglEfektif} />

      {sop.statusVerifikasi === "Revisi" && sop.catatanVerifikasi && (
        <div className="rounded-md border border-red-200 bg-red-50 px-3 py-2.5 mb-3 mt-2">
          <p className="text-xs font-semibold text-red-700 mb-1">Catatan dari Bagian Organisasi:</p>
          <p className="text-xs text-red-700">{sop.catatanVerifikasi}</p>
        </div>
      )}

      <div className="flex items-center gap-2 mt-3">
        <button
          onClick={onEdit}
          className="inline-flex items-center gap-1.5 text-xs font-semibold text-teal-900 bg-white border border-stone-200 rounded-md px-3 py-1.5 hover:border-teal-800 transition-colors"
        >
          <Pencil size={12} strokeWidth={2} />
          Edit
        </button>
        {sop.statusVerifikasi === "Terverifikasi" && (
          <button
            onClick={onRevisi}
            className="inline-flex items-center gap-1.5 text-xs font-semibold text-white bg-teal-800 rounded-md px-3 py-1.5 hover:bg-teal-700 transition-colors"
          >
            <History size={12} strokeWidth={2} />
            Catat Revisi
          </button>
        )}
      </div>
    </div>
  );
}

function SopSayaPage({ opd, onBack, onEdit, onRevisi }) {
  const [list, setList] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = async () => {
    setLoading(true);
    setError("");
    try {
      const data = await apiGet("sop", { id_opd: opd.id_opd });
      setList(data.map(mapSopFromApi));
    } catch (err) {
      setError(err.message || "Gagal memuat data");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
  }, []);

  const menunggu = list.filter((s) => s.statusVerifikasi === "Menunggu").length;
  const revisi = list.filter((s) => s.statusVerifikasi === "Revisi").length;

  return (
    <div className="max-w-4xl mx-auto px-6 py-8">
      <button
        onClick={onBack}
        className="inline-flex items-center gap-1.5 text-sm font-medium text-gray-500 hover:text-teal-800 transition-colors mb-5"
      >
        <ArrowLeft size={15} strokeWidth={2.2} />
        Kembali ke pencarian
      </button>

      <div className="bg-teal-900 rounded-xl px-6 py-6 mb-5">
        <div className="text-xs font-semibold tracking-widest uppercase text-teal-200 mb-2">
          {opd.nama_opd}
        </div>
        <h1
          className="text-2xl leading-snug font-bold text-white"
          style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
        >
          SOP Saya
        </h1>
        {(menunggu > 0 || revisi > 0) && (
          <p className="text-sm text-teal-100 mt-2">
            {menunggu > 0 && <>{menunggu} menunggu verifikasi</>}
            {menunggu > 0 && revisi > 0 && " · "}
            {revisi > 0 && <>{revisi} perlu direvisi</>}
          </p>
        )}
      </div>

      {loading && <p className="text-sm text-gray-500">Memuat…</p>}
      {error && <p className="text-sm text-red-600">{error}</p>}

      {!loading && list.length === 0 && !error && (
        <div className="text-center py-20 border border-dashed border-stone-200 rounded-md">
          <FileText size={28} strokeWidth={1.5} className="mx-auto text-teal-100 mb-3" />
          <p className="text-gray-500 text-sm">Belum ada SOP yang diajukan OPD ini.</p>
        </div>
      )}

      <div className="grid gap-3">
        {list.map((sop) => (
          <SopSayaCard key={sop.id} sop={sop} onEdit={() => onEdit(sop)} onRevisi={() => onRevisi(sop)} />
        ))}
      </div>
    </div>
  );
}

function VerifikasiPage({ verifikatorName, onLogout, onBack, onOpenLog }) {
  const [pending, setPending] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [selectedSop, setSelectedSop] = useState(null);

  const load = async () => {
    setLoading(true);
    setError("");
    try {
      const data = await apiGet("sop", { status_verifikasi: "Menunggu" });
      setPending(data.map(mapSopFromApi));
    } catch (err) {
      setError(err.message || "Gagal memuat data");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
  }, []);

  const handleSelesai = () => {
    setSelectedSop(null);
    load();
  };

  return (
    <div className="min-h-screen bg-stone-50">
      <div className="bg-teal-900">
        <div className="max-w-4xl mx-auto px-6 py-5 flex items-center justify-between gap-4">
          <div>
            <div className="text-xs font-semibold tracking-widest uppercase text-teal-200">
              Bagian Organisasi
            </div>
            <h1 className="text-xl font-bold text-white leading-none mt-1">Verifikasi SOP</h1>
            <p className="text-xs text-teal-200 mt-1">Masuk sebagai {verifikatorName}</p>
          </div>
          <div className="flex flex-col items-end gap-2 shrink-0">
            <div className="flex items-center gap-2">
              <button
                onClick={onOpenLog}
                className="text-xs font-medium text-teal-100 bg-white bg-opacity-10 hover:bg-opacity-20 rounded-full px-3 py-1.5 transition-colors"
              >
                Log Aktivitas
              </button>
              <button
                onClick={onLogout}
                className="text-xs font-medium text-teal-100 bg-white bg-opacity-10 hover:bg-opacity-20 rounded-full px-3 py-1.5 transition-colors"
              >
                Keluar
              </button>
            </div>
            <button
              onClick={onBack}
              className="text-xs font-medium text-teal-100 hover:text-white transition-colors"
            >
              ← Ke portal publik
            </button>
          </div>
        </div>
      </div>

      {selectedSop ? (
        <SopDetailPage
          sop={selectedSop}
          isOperator={false}
          onBack={() => setSelectedSop(null)}
          backLabel="Kembali ke daftar verifikasi"
          verifikasiPanel={{ verifikatorName, onSelesai: handleSelesai }}
        />
      ) : (
        <div className="max-w-4xl mx-auto px-6 py-8">
          <div className="flex items-center justify-between mb-4">
            <p className="text-sm text-gray-500">
              <strong className="text-teal-900">{pending.length}</strong> SOP menunggu verifikasi
            </p>
            <button
              onClick={load}
              className="text-xs font-semibold text-teal-800 border border-teal-800 rounded-md px-3 py-1.5 hover:bg-teal-50 transition-colors"
            >
              Muat ulang
            </button>
          </div>

          {loading && <p className="text-sm text-gray-500">Memuat…</p>}
          {error && <p className="text-sm text-red-600">{error}</p>}

          {!loading && pending.length === 0 && !error && (
            <div className="text-center py-20 border border-dashed border-stone-200 rounded-md">
              <CheckCircle2 size={28} strokeWidth={1.5} className="mx-auto text-teal-100 mb-3" />
              <p className="text-gray-500 text-sm">Tidak ada SOP yang menunggu verifikasi saat ini.</p>
            </div>
          )}

          <div className="grid gap-3">
            {pending.map((sop) => (
              <PendingSopRow key={sop.id} sop={sop} onOpen={() => setSelectedSop(sop)} />
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

const AKSI_BADGE = {
  "Ajukan Baru": { bg: "bg-sky-50", text: "text-sky-700" },
  Edit: { bg: "bg-stone-100", text: "text-stone-600" },
  Revisi: { bg: "bg-amber-50", text: "text-amber-700" },
  Verifikasi: { bg: "bg-emerald-50", text: "text-emerald-700" },
  "Minta Revisi": { bg: "bg-red-50", text: "text-red-700" },
};

function LogAktivitasPage({ onBack }) {
  const [log, setLog] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = async () => {
    setLoading(true);
    setError("");
    try {
      const data = await apiGet("log");
      setLog(data);
    } catch (err) {
      setError(err.message || "Gagal memuat log");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
  }, []);

  return (
    <div className="min-h-screen bg-stone-50">
      <div className="bg-teal-900">
        <div className="max-w-4xl mx-auto px-6 py-5 flex items-center justify-between gap-4">
          <div>
            <div className="text-xs font-semibold tracking-widest uppercase text-teal-200">
              Bagian Organisasi
            </div>
            <h1 className="text-xl font-bold text-white leading-none mt-1">Log Aktivitas</h1>
            <p className="text-xs text-teal-200 mt-1">200 aktivitas terakhir</p>
          </div>
          <button
            onClick={onBack}
            className="text-xs font-medium text-teal-100 bg-white bg-opacity-10 hover:bg-opacity-20 rounded-full px-3 py-1.5 transition-colors shrink-0"
          >
            ← Kembali
          </button>
        </div>
      </div>

      <div className="max-w-4xl mx-auto px-6 py-8">
        <div className="flex items-center justify-between mb-4">
          <p className="text-sm text-gray-500">
            <strong className="text-teal-900">{log.length}</strong> aktivitas tercatat
          </p>
          <button
            onClick={load}
            className="text-xs font-semibold text-teal-800 border border-teal-800 rounded-md px-3 py-1.5 hover:bg-teal-50 transition-colors"
          >
            Muat ulang
          </button>
        </div>

        {loading && <p className="text-sm text-gray-500">Memuat…</p>}
        {error && <p className="text-sm text-red-600">{error}</p>}

        {!loading && log.length === 0 && !error && (
          <div className="text-center py-20 border border-dashed border-stone-200 rounded-md">
            <History size={28} strokeWidth={1.5} className="mx-auto text-teal-100 mb-3" />
            <p className="text-gray-500 text-sm">Belum ada aktivitas tercatat.</p>
          </div>
        )}

        <div className="grid gap-2">
          {log.map((l, i) => {
            const badge = AKSI_BADGE[l.aksi] || { bg: "bg-stone-100", text: "text-stone-600" };
            return (
              <div
                key={i}
                className="bg-white border border-stone-200 rounded-xl px-5 py-4 flex items-start justify-between gap-4"
              >
                <div className="min-w-0">
                  <div className="flex items-center gap-2 mb-1 flex-wrap">
                    <span className={`text-xs font-semibold rounded-full px-2.5 py-1 ${badge.bg} ${badge.text}`}>
                      {l.aksi}
                    </span>
                    <span className="text-xs text-gray-500">oleh {l.aktor}</span>
                  </div>
                  <p className="text-sm font-medium text-teal-900 truncate">
                    {l.nomor_sop ? `${l.nomor_sop} · ` : ""}
                    {l.judul_sop || "-"}
                  </p>
                  {l.keterangan && <p className="text-xs text-gray-500 mt-1">{l.keterangan}</p>}
                </div>
                <span className="text-xs text-stone-400 shrink-0 whitespace-nowrap">{l.waktu}</span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}


function SopFormPage({ mode, initialSop, opd, onBack }) {
  const [submitted, setSubmitted] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState("");
  const formRef = useRef(null);

  const eyebrow = mode === "tambah" ? "Pengajuan OPD" : mode === "revisi" ? "Catat Revisi" : "Edit Data";
  const title = mode === "tambah" ? "Ajukan SOP Baru" : mode === "revisi" ? "Catat Revisi SOP" : "Edit SOP";
  const successText =
    mode === "tambah"
      ? "SOP akan tayang di portal publik setelah diverifikasi oleh Bagian Organisasi."
      : "Perubahan akan tayang di portal publik setelah diverifikasi oleh Bagian Organisasi.";
  const submitLabel = mode === "tambah" ? "Kirim untuk Verifikasi" : "Simpan Perubahan";

  const handleSubmit = async () => {
    const fd = new FormData(formRef.current);
    const get = (name) => (fd.get(name) || "").toString().trim();

    if (mode === "revisi" && (!get("tgl_revisi") || !get("keterangan"))) {
      setError("Tanggal revisi dan keterangan wajib diisi");
      return;
    }
    setError("");
    setSubmitting(true);
    try {
      if (mode === "tambah") {
        await apiPost("submitSop", {
          id_opd: opd.id_opd,
          nomor_sop: get("nomor_sop"),
          judul_sop: get("judul_sop"),
          bidang_bagian: get("bidang_bagian"),
          seksi_subbid_subbag: get("seksi_subbid_subbag"),
          abstrak: get("abstrak"),
          tgl_pembuatan: get("tgl_pembuatan"),
          tgl_efektif: get("tgl_efektif"),
          status: get("status"),
          disahkan_oleh: get("disahkan_oleh"),
          link_drive: get("link_drive"),
        });
      } else if (mode === "edit") {
        await apiPost("editSop", {
          id_sop: initialSop.id,
          id_opd: opd.id_opd,
          nomor_sop: get("nomor_sop"),
          judul_sop: get("judul_sop"),
          bidang_bagian: get("bidang_bagian"),
          seksi_subbid_subbag: get("seksi_subbid_subbag"),
          abstrak: get("abstrak"),
          tgl_pembuatan: get("tgl_pembuatan"),
          tgl_efektif: get("tgl_efektif"),
          status: get("status"),
          disahkan_oleh: get("disahkan_oleh"),
          link_drive: get("link_drive"),
        });
      } else {
        await apiPost("revisiSop", {
          id_sop: initialSop.id,
          id_opd: opd.id_opd,
          tgl_revisi: get("tgl_revisi"),
          keterangan: get("keterangan"),
          link_drive: get("link_drive"),
          tgl_efektif: get("tgl_efektif"),
        });
      }
      setSubmitted(true);
    } catch (err) {
      setError(err.message || "Gagal mengirim data, coba lagi.");
    } finally {
      setSubmitting(false);
    }
  };

  if (submitted) {
    return (
      <div className="max-w-4xl mx-auto px-6 py-8">
        <div className="bg-white border border-stone-200 rounded-xl px-6 py-16 text-center">
          <CheckCircle2 size={40} strokeWidth={1.6} className="mx-auto text-green-700 mb-3" />
          <p className="text-base font-semibold text-teal-900 mb-1">
            {mode === "tambah" ? "Pengajuan terkirim" : "Perubahan tersimpan"}
          </p>
          <p className="text-sm text-gray-500 max-w-sm mx-auto mb-6">{successText}</p>
          <button
            onClick={onBack}
            className="text-sm font-semibold text-teal-900 border border-teal-800 rounded-md px-5 py-2.5 hover:bg-teal-700 hover:text-white transition-colors"
          >
            Kembali ke pencarian
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-4xl mx-auto px-6 py-8">
      <button
        onClick={onBack}
        className="inline-flex items-center gap-1.5 text-sm font-medium text-gray-500 hover:text-teal-800 transition-colors mb-5"
      >
        <ArrowLeft size={15} strokeWidth={2.2} />
        Kembali ke pencarian
      </button>

      <div className="bg-teal-900 rounded-xl px-6 py-6 mb-5">
        <div className="text-xs font-semibold tracking-widest uppercase text-teal-200 mb-2">{eyebrow}</div>
        <h1
          className="text-2xl leading-snug font-bold text-white"
          style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
        >
          {title}
        </h1>
      </div>

      <form ref={formRef} className="bg-white border border-stone-200 rounded-xl px-6 py-6 space-y-4">
        <div className="flex items-center gap-2 rounded-md border border-stone-200 bg-teal-50 px-3 py-2.5">
          <Building2 size={15} strokeWidth={2.2} className="text-teal-700" />
          <span className="text-sm font-semibold text-teal-900">{opd.nama_opd}</span>
          <span className="ml-auto text-xs text-gray-500">OPD Anda</span>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">Nomor SOP</label>
            <input
              name="nomor_sop"
              required
              type="text"
              defaultValue={initialSop?.nomor || ""}
              placeholder="cth. 1.4.1.1.1"
              className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm font-mono text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">Disahkan oleh</label>
            <input
              name="disahkan_oleh"
              required
              type="text"
              defaultValue={initialSop?.disahkanOleh || ""}
              placeholder="cth. Kepala Dinas…"
              className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
          </div>
        </div>

        <div>
          <label className="block text-xs font-semibold text-gray-700 mb-1.5">Judul SOP</label>
          <input
            name="judul_sop"
            required
            type="text"
            defaultValue={initialSop?.judul || ""}
            placeholder="cth. Penyusunan Rencana Kerja…"
            className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
          />
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">Bidang / Bagian</label>
            <input
              name="bidang_bagian"
              required
              type="text"
              defaultValue={initialSop?.bidang || ""}
              placeholder="cth. Bidang Pelayanan Kesehatan"
              className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">Seksi / Subbid / Subbag</label>
            <input
              name="seksi_subbid_subbag"
              required
              type="text"
              defaultValue={initialSop?.seksi || ""}
              placeholder="cth. Seksi Pelayanan Kesehatan Primer"
              className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
          </div>
        </div>

        <div>
          <label className="block text-xs font-semibold text-gray-700 mb-1.5">
            Abstraksi <span className="text-stone-400 font-normal">(opsional)</span>
          </label>
          <textarea
            name="abstrak"
            rows={3}
            defaultValue={initialSop?.abstrak || ""}
            placeholder="Ringkasan singkat isi dan tujuan SOP ini…"
            className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800 resize-none"
          />
        </div>

        <div>
          <label className="block text-xs font-semibold text-gray-700 mb-1.5">Status</label>
          <select
            name="status"
            required
            defaultValue={initialSop?.status || "Berlaku"}
            className="w-full sm:w-64 rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
          >
            <option value="Berlaku">Berlaku</option>
            <option value="Tidak Berlaku">Tidak Berlaku</option>
          </select>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">Tgl Pembuatan</label>
            <input
              name="tgl_pembuatan"
              required
              type="date"
              defaultValue={toIsoDate(initialSop?.tglPembuatan)}
              className="w-full rounded-md border border-stone-200 bg-white px-2.5 py-2.5 text-sm text-teal-900 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
          </div>
          <div>
            <label
              className={`block text-xs font-semibold mb-1.5 ${
                mode === "revisi" ? "text-teal-800" : "text-gray-700"
              }`}
            >
              Tgl Revisi {mode === "revisi" && <span className="text-teal-700">*</span>}
            </label>
            <input
              name="tgl_revisi"
              required={mode === "revisi"}
              autoFocus={mode === "revisi"}
              type="date"
              defaultValue={mode === "revisi" ? "" : toIsoDate(initialSop?.tglRevisi)}
              className={`w-full rounded-md bg-white px-2.5 py-2.5 text-sm text-teal-900 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800 ${
                mode === "revisi" ? "border-2 border-amber-500" : "border border-stone-200"
              }`}
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-gray-700 mb-1.5">Tgl Efektif</label>
            <input
              name="tgl_efektif"
              required
              type="date"
              defaultValue={toIsoDate(initialSop?.tglEfektif)}
              className="w-full rounded-md border border-stone-200 bg-white px-2.5 py-2.5 text-sm text-teal-900 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
          </div>
        </div>

        {mode === "revisi" && (
          <div>
            <label className="block text-xs font-semibold text-teal-800 mb-1.5">
              Keterangan Revisi <span className="text-teal-700">*</span>
            </label>
            <textarea
              name="keterangan"
              required
              rows={2}
              placeholder="cth. Penyesuaian alur paraf berjenjang…"
              className="w-full rounded-md border-2 border-amber-500 bg-white px-3 py-2.5 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 resize-none"
            />
            <p className="mt-1.5 text-xs text-stone-400">Ringkasan singkat apa yang berubah dari versi sebelumnya.</p>
          </div>
        )}

        <div>
          <label className="block text-xs font-semibold text-gray-700 mb-1.5">Link Google Drive (PDF SOP)</label>
          <div className="relative">
            <Link2 size={15} strokeWidth={2} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
            <input
              name="link_drive"
              required
              type="url"
              placeholder="https://drive.google.com/…"
              className="w-full pl-9 pr-3 py-2.5 rounded-md border border-stone-200 bg-white text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800"
            />
          </div>
          <p className="mt-1.5 text-xs text-stone-400">
            {mode === "revisi"
              ? "Unggah versi PDF SOP terbaru dan tempel link-nya di sini."
              : 'Pastikan akses berbagi link diatur ke "Siapa saja yang memiliki link".'}
          </p>
        </div>

        {error && (
          <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-md px-3 py-2">{error}</p>
        )}

        <div className="pt-2 flex items-center justify-end gap-2 border-t border-stone-200">
          <button
            type="button"
            onClick={onBack}
            className="text-sm font-medium text-gray-500 px-4 py-2.5 hover:text-gray-700 transition-colors"
          >
            Batal
          </button>
          <button
            type="button"
            onClick={handleSubmit}
            disabled={submitting}
            className="inline-flex items-center gap-2 text-sm font-semibold text-white bg-teal-800 rounded-md px-4 py-2.5 hover:bg-teal-700 transition-colors disabled:opacity-50"
          >
            <Send size={14} strokeWidth={2.2} />
            {submitting ? "Mengirim…" : submitLabel}
          </button>
        </div>
      </form>
    </div>
  );
}

export default function SopPortal() {
  const [query, setQuery] = useState("");
  const [opdFilter, setOpdFilter] = useState("");
  const [onlyPerluDitinjau, setOnlyPerluDitinjau] = useState(false);
  const [statusFilter, setStatusFilter] = useState("");
  const [tglDari, setTglDari] = useState("");
  const [tglSampai, setTglSampai] = useState("");
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [opdOpen, setOpdOpen] = useState(false);
  const [modalMode, setModalMode] = useState("tambah");
  const [activeSop, setActiveSop] = useState(null);
  // OPD yang sedang login sebagai operator: { id_opd, nama_opd, singkatan } | null (publik)
  const [loggedInOpd, setLoggedInOpd] = useState(null);
  const [showLogin, setShowLogin] = useState(false);
  // "list" | "detail" | "form" | "verifikasi"
  const [view, setView] = useState("list");
  // Halaman tujuan tombol "Kembali" di SopFormPage: "list" | "detail" | "sopsaya"
  const [formOrigin, setFormOrigin] = useState("list");
  const [detailSop, setDetailSop] = useState(null);

  // Verifikator (Bagian Organisasi)
  const [verifikatorName, setVerifikatorName] = useState(null);
  const [showVerifikatorLogin, setShowVerifikatorLogin] = useState(false);

  // ---- Data dari backend Apps Script ----
  const [opdList, setOpdList] = useState([]);
  const [sopList, setSopList] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState("");

  const loadData = async () => {
    setLoading(true);
    setLoadError("");
    try {
      const [opdData, sopData] = await Promise.all([apiGet("opd"), apiGet("sop")]);
      setOpdList(opdData);
      setSopList(sopData.map(mapSopFromApi));
    } catch (err) {
      setLoadError(err.message || "Gagal memuat data dari server");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, []);

  // Refresh data setiap kali berpindah kembali ke halaman utama/publik,
  // supaya data yang tampil selalu terbaru (mis. habis verifikasi/ajukan SOP)
  const isFirstView = useRef(true);
  useEffect(() => {
    if (isFirstView.current) {
      isFirstView.current = false;
      return;
    }
    if (view === "list") {
      loadData();
    }
  }, [view]);

  const openDetail = (sop) => {
    setDetailSop(sop);
    setView("detail");
  };
  const closeDetail = () => {
    setView("list");
  };

  const openTambah = () => {
    setModalMode("tambah");
    setActiveSop(null);
    setFormOrigin("list");
    setView("form");
  };
  const openEdit = (sop, origin = "detail") => {
    setModalMode("edit");
    setActiveSop(sop);
    setDetailSop(sop);
    setFormOrigin(origin);
    setView("form");
  };
  const openRevisi = (sop, origin = "detail") => {
    setModalMode("revisi");
    setActiveSop(sop);
    setDetailSop(sop);
    setFormOrigin(origin);
    setView("form");
  };
  const handleLogin = (opd) => {
    setLoggedInOpd(opd); // { id_opd, nama_opd, singkatan }
    setShowLogin(false);
  };
  const handleLogout = () => {
    setLoggedInOpd(null);
    if (view !== "list") setView("list");
  };
  const handleVerifikatorLogin = (nama) => {
    setVerifikatorName(nama);
    setShowVerifikatorLogin(false);
    setView("verifikasi");
  };
  const handleVerifikatorLogout = () => {
    setVerifikatorName(null);
    setView("list");
  };

  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    return sopList.filter((s) => {
      const matchesQuery =
        q === "" ||
        s.judul.toLowerCase().includes(q) ||
        s.nomor.toLowerCase().includes(q) ||
        (s.bidang || "").toLowerCase().includes(q) ||
        (s.seksi || "").toLowerCase().includes(q);
      const matchesOpd = opdFilter === "" || s.opd === opdFilter;
      const matchesReview = !onlyPerluDitinjau || perluDitinjau(s.tglEfektif);
      const matchesStatus = statusFilter === "" || s.status === statusFilter;
      const matchesDari = tglDari === "" || (s.tglEfektif && s.tglEfektif >= tglDari);
      const matchesSampai = tglSampai === "" || (s.tglEfektif && s.tglEfektif <= tglSampai);
      return (
        matchesQuery && matchesOpd && matchesReview && matchesStatus && matchesDari && matchesSampai
      );
    });
  }, [sopList, query, opdFilter, onlyPerluDitinjau, statusFilter, tglDari, tglSampai]);

  const PAGE_SIZE = 15;
  const [page, setPage] = useState(1);
  useEffect(() => {
    setPage(1);
  }, [query, opdFilter, sopList, onlyPerluDitinjau, statusFilter, tglDari, tglSampai]);
  const totalPages = Math.max(1, Math.ceil(results.length / PAGE_SIZE));
  const pagedResults = useMemo(() => {
    const start = (page - 1) * PAGE_SIZE;
    return results.slice(start, start + PAGE_SIZE);
  }, [results, page]);

  const latest = useMemo(() => {
    return [...sopList].sort((a, b) => (a.diverifikasi < b.diverifikasi ? 1 : -1)).slice(0, 5);
  }, [sopList]);

  const opdCounts = useMemo(() => {
    const map = {};
    sopList.forEach((s) => {
      map[s.opd] = (map[s.opd] || 0) + 1;
    });
    return Object.entries(map)
      .map(([opd, count]) => ({ opd, count }))
      .sort((a, b) => b.count - a.count);
  }, [sopList]);

  const perluDitinjauCount = useMemo(
    () => sopList.filter((s) => perluDitinjau(s.tglEfektif)).length,
    [sopList]
  );

  const stats = useMemo(() => {
    return [
      { label: "SOP Terverifikasi", value: sopList.length, icon: ClipboardList, color: "emerald" },
      { label: "OPD Berpartisipasi", value: opdCounts.length, icon: Landmark, color: "blue" },
      { label: "Total OPD di Inhu", value: opdList.length, icon: Layers, color: "slate" },
      { label: "Perlu Ditinjau", value: perluDitinjauCount, icon: History, color: "amber" },
    ];
  }, [sopList, opdList, opdCounts, perluDitinjauCount]);

  if (view === "verifikasi" && verifikatorName) {
    return (
      <VerifikasiPage
        verifikatorName={verifikatorName}
        onLogout={handleVerifikatorLogout}
        onBack={() => setView("list")}
        onOpenLog={() => setView("log")}
      />
    );
  }

  if (view === "log" && verifikatorName) {
    return <LogAktivitasPage onBack={() => setView("verifikasi")} />;
  }

  return (
    <div className="min-h-screen bg-stone-50 font-sans" style={{ fontFamily: "'Inter', system-ui, sans-serif" }}>
      {/* Header — teal bar with logo lockup, matching Patuhdiri/MANDALA branding */}
      <div className="bg-teal-900">
        <div className="max-w-4xl mx-auto px-6 py-5 flex items-center justify-between gap-4">
          <div className="flex items-start gap-3">
            <img
              src="/logo-portal-sop.png"
              alt="Logo Portal SOP"
              className="shrink-0 h-24 w-auto object-contain"
            />
            <div>
              <h1 className="text-xl font-bold text-amber-400 leading-none">Portal SOP</h1>
              <p className="text-sm italic text-teal-100 mt-1">Registri Standar Operasional Prosedur</p>
              <p className="text-xs text-teal-300 mt-0.5">
                Bagian Organisasi Sekretariat Daerah Kabupaten Indragiri Hulu
              </p>
            </div>
          </div>
          {loggedInOpd ? (
            <div className="flex flex-col items-end gap-2 shrink-0">
              <div className="flex items-center gap-2 flex-wrap justify-end">
                <button
                  onClick={handleLogout}
                  className="text-xs font-medium text-teal-100 bg-white bg-opacity-10 hover:bg-opacity-20 rounded-full px-3 py-1.5 whitespace-nowrap transition-colors"
                >
                  Keluar
                </button>
              </div>
              <span className="inline-flex items-center gap-1.5 text-xs font-medium text-teal-900 bg-amber-400 rounded-full px-3 py-1.5 whitespace-nowrap">
                <ShieldCheck size={13} strokeWidth={2.3} />
                {loggedInOpd.nama_opd}
              </span>
            </div>
          ) : (
            <button
              onClick={() => setShowLogin(true)}
              className="inline-flex items-center gap-1.5 text-xs font-medium text-teal-100 bg-white bg-opacity-10 hover:bg-opacity-20 rounded-full px-3 py-1.5 whitespace-nowrap transition-colors shrink-0"
            >
              <ShieldCheck size={13} strokeWidth={2.3} />
              Masuk
            </button>
          )}
        </div>
      </div>

      {(loading || loadError) && (
        <div
          className={`max-w-4xl mx-auto px-6 py-3 text-sm flex items-center justify-between gap-3 ${
            loadError ? "text-red-700" : "text-gray-500"
          }`}
        >
          <span>{loadError ? `Gagal memuat data: ${loadError}` : "Memuat data SOP…"}</span>
          {loadError && (
            <button
              onClick={loadData}
              className="text-xs font-semibold text-teal-800 border border-teal-800 rounded-md px-3 py-1.5 hover:bg-teal-50 transition-colors shrink-0"
            >
              Coba lagi
            </button>
          )}
        </div>
      )}

      {view === "list" && (
      <header className="relative border-b border-stone-200 bg-teal-900 min-h-[360px]">
        <div className="absolute inset-0 overflow-hidden">
          {/* Background hero — foto Menara Air (rasio panjang, sudah menyatu dengan teal) */}
          <img
            src="data:image/jpeg;base64,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M8+of2bJOt7i/TzHkzFjzGIzu4Yj+degj4Maf42W3gtNWTRdSvZnW1SePaskic+Uxz8hJyATgEetanhrxtpWi6HfWVxeOfEWrY06BrW3SaWGIZLEBsAZwEwPXNd78AtG0z4xeLtKtdV1yHULSC0AfStSsTDNeRqWAKSKeWQ8g5yBXNUk43fY6IRT0PX/ANlDwv4u8HaFLYeJrLUX2/ulkuLeGCC3C/3dvzuTxya+fv8Agof4yvPB/wATPDX2bxDqEdzCi31tY2rFFt25USZ9SR9etfWH7RHxpn+B/g7T4tB0/wDtPWMAW+nTRySCWBAQfnHAIwDye1flx8Vfizqvx/8Ai0fEfiix8iNlFt9lhYokaIDtUE9+59a5cPBzqe0aOirJQhyI6eiiivsz8UCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAClXqKSlX7woAeByPpTqT+L8KWgAooooAd2FFIKWgBRj0qQd6jp4bPtQAucdqcPu9Kbn3px6DntQajox7U8daYhGetPFAC5pR/Wk70q0APA+UcetHQdOwoB6c80o5B57VL2AcGK9AetKpyckHOKQgnAyKeMgcmoAcDSjH600detKAfXvQUidQNg47VKOp4z71Eudgwe1PGcjnvSauMkA56Y607Hyg44ximrnPUd6eASuAR0rMB4wT9KVRg0mM4py9aLFLcWnJ978KbTl+9SKTHjgn1qa2z5jfSq4JLelWLXO4544/rQUty1SjrRQRzxQaj1ABPOacBzTUGBT15YUAKpIOD0qQUhANKBQaIUZ71De2MWpWz282fLfrtOKmAy2fwpwyOaBnNeItAlvpdMtLePbZxZLHsMdM1kaD4JnmvWNxCYgHXJYfeXBJIr0Bc7lz6j+dXwfWk2JI8l8cWc0WtzSiBora0iVEz0bPAI/E1zVvCbx3ha5VYoUL5buT2Fe1+KbE3mhXiRRJJO0fGRk8HNeI6baxXF08csvllsjp0PvTQy7FqMbxWmmxorxrkySBMszZ7fhWdav5V1LG2EjkBBDDI9vp9al03TZZZ3ePJMWSyqOQoHX6U8WYu7RcAJOhyVIwdp6H9R+dAH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cLctulEqjBLHAzuxnNdnc3UTsu24aNSjZXaD+Of6e9eLL3XZnpp31KksExjbNvG7bVzjvUckcyucWqYMmc47Y61bd7faSbmQ5jX7q4wKbIbfeymZ9xlxyO+KzND54/bdvNOsP2dtck1eaeykBH2SC1maL7TNn5UbbyV6nHGcV+SF9r0qQvBFHDaSOdxW1jCmMejN3PtX7CftH/s8Wn7QGh6dZDXLmwmtHZrcFiI0JceZLtH3nCAhc8c1+Yv7RPwTu/ht8W9S8NWOlzW1goD2srhiJolA/eFjwc4JJ6A16mFlHlt1PPxMZN3Wxeooor64/GgooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigApV+8KSlX7woAeOv4U6mgYb8KdQAUUUUAL0pRSUtACjg08EkVHT0oAXnsKeASBwKauM9+9KMYoNSRM+gFPqND9etSAdKAFB4/CpE4HQHmou1SqfX1oAD347UYIpQAc9elOwM/lQAgJx0zT+cdB0pAATxng0uOe/Sk1cBw69KUU0DmnAYNQ1YCcfcHAzjpUiqM1GvCA+1SpnnjjtSLHcgj0qQ9M9D61Go5Geeal4I9qAFXOelOpF60tZsA6ewp6sATz3xTQNxwakVemeaQDlGSMVPD98/Sq6qQwwx4/WrEP3z9KbNIFilCAsGwKSlXqKk3JRyeaeDnAFMA4zSx96AHBBnPc04fLnJyKSlC5+lADVXecnpUyrx7etNFSggjiixTYqAb1H+0P51e6VRT76f7w/nWiv3uaTGth6fd561InT2plPUFR7Uih4FLSLyKUNzyDj1oKJFx16VLJ0FRxf1qagYdqcvOBTacM8UANOg2ky3MyQpHd3CeW8+35iK1o4lEIjIBULtwRweMVHa5MC/jUwFI6I7HmPxi8DXOuWtte2IZzaoIo7ONOOpJP5Ck074e3h1+0u0iVYfssLuso+VhgB4z/ADr1JB0OeakX0qXIdhyRLDGiRqFRRgKOgFedfHm387wWC0kMcazAsXGXJ7Bfc16SDxXnnxj8LXHiTTreRbiO2srNXlcu2N74wige9Edwex80i5ddqIFQryH2gMfbNW9Ue3CQ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"
            alt=""
            className="w-full h-full object-contain object-right"
          />
        </div>

        <div className="relative max-w-4xl mx-auto px-6 pt-10 pb-8">
          <h1
            className="text-3xl sm:text-4xl leading-none font-semibold text-white mb-2"
            style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
          >
            Registri Standar Operasional Prosedur
          </h1>
          <p className="text-teal-100 text-sm max-w-xl mb-6">
            Cari, lihat, dan unduh dokumen SOP dari seluruh perangkat daerah di lingkungan
            Kabupaten Indragiri Hulu yang telah diverifikasi Bagian Organisasi.
          </p>

          {/* Search bar — glassy */}
          <div className="relative w-3/4">
            <Search
              size={17}
              strokeWidth={2.2}
              className="absolute z-10 left-4 top-1/2 -translate-y-1/2 text-white"
            />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              type="text"
              placeholder="Cari judul atau nomor SOP…"
              className="w-full pl-11 pr-4 py-3.5 rounded-md border border-white border-opacity-20 bg-white bg-opacity-10 backdrop-blur-md text-sm text-white placeholder-white placeholder-opacity-70 focus:outline-none focus:ring-2 focus:ring-white focus:ring-opacity-40 focus:border-white focus:border-opacity-40 transition-all"
            />
          </div>

          {/* OPD filter dropdown + Ajukan SOP Baru */}
          <div className="mt-3 grid grid-cols-2 gap-3 items-start">
            <div className="flex items-center gap-2 flex-wrap">
              <div className="relative w-full">
                <button
                  onClick={() => setOpdOpen((v) => !v)}
                  className="w-full inline-flex items-center gap-2 text-sm font-medium text-white border border-white border-opacity-20 rounded-md pl-3 pr-2.5 py-2 bg-white bg-opacity-10 backdrop-blur-md hover:bg-opacity-20 transition-colors"
                >
                  <Building2 size={14} strokeWidth={2.2} className="text-white text-opacity-70 shrink-0" />
                  <span className="truncate">{opdFilter || "Semua OPD"}</span>
                  <ChevronDown size={14} strokeWidth={2.2} className="text-white text-opacity-70 ml-auto shrink-0" />
                </button>
                {opdOpen && (
                  <div className="absolute z-10 mt-1.5 w-full bg-white border border-stone-200 rounded-md shadow-lg py-1.5 max-h-64 overflow-auto">
                    <button
                      onClick={() => {
                        setOpdFilter("");
                        setOpdOpen(false);
                      }}
                      className="w-full text-left px-3.5 py-2 text-sm text-gray-700 hover:bg-stone-50"
                    >
                      Semua OPD
                    </button>
                    {opdList.map((o) => (
                      <button
                        key={o.id_opd}
                        onClick={() => {
                          setOpdFilter(o.nama_opd);
                          setOpdOpen(false);
                        }}
                        className="w-full text-left px-3.5 py-2 text-sm text-gray-700 hover:bg-stone-50"
                      >
                        {o.nama_opd}
                      </button>
                    ))}
                  </div>
                )}
              </div>

              {opdFilter && (
                <Chip onRemove={() => setOpdFilter("")}>{opdFilter}</Chip>
              )}
              <button
                onClick={() => setShowAdvanced((v) => !v)}
                className="inline-flex items-center gap-1.5 text-xs font-medium text-teal-100 hover:text-white underline transition-colors"
              >
                Filter lanjutan {showAdvanced ? "▲" : "▼"}
              </button>
            </div>

            <div className="flex justify-end gap-2">
              {loggedInOpd && (
                <>
                  <button
                    onClick={() => setView("sopsaya")}
                    className="inline-flex items-center gap-2 text-sm font-medium text-white border border-white border-opacity-20 rounded-md px-4 py-2 bg-white bg-opacity-10 backdrop-blur-md hover:bg-opacity-20 transition-colors"
                  >
                    SOP Saya
                  </button>
                  <button
                    onClick={openTambah}
                    className="inline-flex items-center gap-2 bg-teal-800 text-white font-semibold text-sm px-4 py-2 rounded-md hover:bg-teal-700 transition-colors"
                  >
                    <Plus size={16} />
                    Ajukan SOP Baru
                  </button>
                </>
              )}
            </div>

            {showAdvanced && (
              <div className="col-span-2 grid grid-cols-1 sm:grid-cols-3 gap-3 mt-1 p-3 rounded-md border border-white border-opacity-20 bg-white bg-opacity-10 backdrop-blur-md">
                <div>
                  <label className="block text-xs font-medium text-teal-100 mb-1.5">Status</label>
                  <select
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value)}
                    className="w-full rounded-md border border-white border-opacity-20 bg-white bg-opacity-10 px-2.5 py-2 text-sm text-white focus:outline-none focus:ring-2 focus:ring-white focus:ring-opacity-40"
                  >
                    <option value="" className="text-teal-900">Semua Status</option>
                    <option value="Berlaku" className="text-teal-900">Berlaku</option>
                    <option value="Tidak Berlaku" className="text-teal-900">Tidak Berlaku</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-teal-100 mb-1.5">Efektif dari</label>
                  <input
                    type="date"
                    value={tglDari}
                    onChange={(e) => setTglDari(e.target.value)}
                    className="w-full rounded-md border border-white border-opacity-20 bg-white bg-opacity-10 px-2.5 py-2 text-sm text-white focus:outline-none focus:ring-2 focus:ring-white focus:ring-opacity-40"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-teal-100 mb-1.5">Efektif sampai</label>
                  <input
                    type="date"
                    value={tglSampai}
                    onChange={(e) => setTglSampai(e.target.value)}
                    className="w-full rounded-md border border-white border-opacity-20 bg-white bg-opacity-10 px-2.5 py-2 text-sm text-white focus:outline-none focus:ring-2 focus:ring-white focus:ring-opacity-40"
                  />
                </div>
              </div>
            )}
          </div>
        </div>
      </header>
      )}

      {view === "detail" && detailSop ? (
        <SopDetailPage
          sop={detailSop}
          isOperator={loggedInOpd?.id_opd === detailSop.idOpd}
          onBack={closeDetail}
          onEdit={() => openEdit(detailSop)}
          onRevisi={() => openRevisi(detailSop)}
        />
      ) : view === "form" && loggedInOpd ? (
        <SopFormPage
          mode={modalMode}
          initialSop={activeSop}
          opd={loggedInOpd}
          onBack={() => setView(formOrigin)}
        />
      ) : view === "sopsaya" && loggedInOpd ? (
        <SopSayaPage
          opd={loggedInOpd}
          onBack={() => setView("list")}
          onEdit={(sop) => openEdit(sop, "sopsaya")}
          onRevisi={(sop) => openRevisi(sop, "sopsaya")}
        />
      ) : view === "opdlist" ? (
        <OpdListPage
          opdCounts={opdCounts}
          onBack={() => setView("list")}
          onSelectOpd={(opd) => {
            setOpdFilter(opd);
            setView("list");
          }}
        />
      ) : (
        <>
          {/* Dashboard stats */}
          <section className="max-w-4xl mx-auto px-6 pt-8">
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 sm:gap-3">
              {stats.map((s) => (
                <StatCard
                  key={s.label}
                  {...s}
                  onClick={
                    s.label === "OPD Berpartisipasi"
                      ? () => setView("opdlist")
                      : s.label === "Perlu Ditinjau"
                      ? () => setOnlyPerluDitinjau(true)
                      : undefined
                  }
                />
              ))}
            </div>
          </section>

          {/* SOP Terbaru */}
          <section className="max-w-4xl mx-auto px-6 pt-8">
            <div className="flex items-center gap-2 mb-3">
              <Sparkles size={15} strokeWidth={2.2} className="text-teal-700" />
              <h2 className="text-sm font-semibold text-teal-900">SOP Terbaru Diverifikasi</h2>
            </div>
            <div className="flex gap-3 overflow-x-auto pb-2 snap-x -mx-6 px-6">
              {latest.map((sop) => (
                <LatestSopCard key={sop.id} sop={sop} onOpen={() => openDetail(sop)} />
              ))}
            </div>
          </section>

          {/* Results */}
          <main className="max-w-4xl mx-auto px-6 py-8 print:hidden">
            <div className="flex items-center justify-between gap-2 mb-3 pt-2 border-t border-stone-200 flex-wrap">
              <div className="flex items-center gap-2">
                <ClipboardList size={15} strokeWidth={2.2} className="text-teal-900" />
                <h2 className="text-sm font-semibold text-teal-900">Jelajahi Semua SOP</h2>
              </div>
              {results.length > 0 && (
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => exportExcel(results)}
                    className="inline-flex items-center gap-1.5 text-xs font-semibold text-teal-900 border border-stone-200 rounded-md px-3 py-1.5 hover:border-teal-800 transition-colors"
                  >
                    <FileSpreadsheet size={13} strokeWidth={2.2} />
                    Unduh Excel
                  </button>
                  <button
                    onClick={() => window.print()}
                    className="inline-flex items-center gap-1.5 text-xs font-semibold text-teal-900 border border-stone-200 rounded-md px-3 py-1.5 hover:border-teal-800 transition-colors"
                  >
                    <Printer size={13} strokeWidth={2.2} />
                    Cetak PDF
                  </button>
                </div>
              )}
            </div>
            <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
              <div className="flex items-center gap-2 text-sm text-gray-500">
                <CalendarCheck size={14} strokeWidth={2.2} className="text-green-700" />
                <span>
                  <strong className="text-teal-900">{results.length}</strong> SOP terverifikasi ditemukan
                </span>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                {onlyPerluDitinjau && (
                  <Chip onRemove={() => setOnlyPerluDitinjau(false)}>Perlu ditinjau</Chip>
                )}
                {statusFilter && (
                  <Chip onRemove={() => setStatusFilter("")}>{statusFilter}</Chip>
                )}
                {tglDari && (
                  <Chip onRemove={() => setTglDari("")}>Dari {formatTanggal(tglDari)}</Chip>
                )}
                {tglSampai && (
                  <Chip onRemove={() => setTglSampai("")}>Sampai {formatTanggal(tglSampai)}</Chip>
                )}
              </div>
            </div>

            {results.length === 0 ? (
              <div className="text-center py-20 border border-dashed border-stone-200 rounded-md">
                <FileText size={28} strokeWidth={1.5} className="mx-auto text-teal-100 mb-3" />
                <p className="text-gray-500 text-sm mb-1">Tidak ada SOP yang cocok dengan pencarian.</p>
                <p className="text-stone-400 text-xs">Coba kata kunci lain atau ubah filter OPD.</p>
              </div>
            ) : (
              <>
                <div className="grid gap-3">
                  {pagedResults.map((sop) => (
                    <SopCard
                      key={sop.id}
                      sop={sop}
                      onOpen={() => openDetail(sop)}
                      onEdit={() => openEdit(sop)}
                      onRevisi={() => openRevisi(sop)}
                      canManage={loggedInOpd?.id_opd === sop.idOpd}
                    />
                  ))}
                </div>

                {totalPages > 1 && (
                  <div className="flex items-center justify-center gap-1.5 mt-6">
                    <button
                      onClick={() => setPage((p) => Math.max(1, p - 1))}
                      disabled={page === 1}
                      className="text-sm font-medium text-gray-700 border border-stone-200 rounded-md px-3 py-1.5 hover:border-teal-800 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
                    >
                      ← Sebelumnya
                    </button>
                    {Array.from({ length: totalPages }, (_, i) => i + 1).map((n) => (
                      <button
                        key={n}
                        onClick={() => setPage(n)}
                        className={`w-8 h-8 text-sm font-medium rounded-md transition-colors ${
                          n === page
                            ? "bg-teal-800 text-white"
                            : "text-gray-700 border border-stone-200 hover:border-teal-800"
                        }`}
                      >
                        {n}
                      </button>
                    ))}
                    <button
                      onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
                      disabled={page === totalPages}
                      className="text-sm font-medium text-gray-700 border border-stone-200 rounded-md px-3 py-1.5 hover:border-teal-800 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
                    >
                      Selanjutnya →
                    </button>
                  </div>
                )}
              </>
            )}
          </main>

          {/* Laporan cetak — cuma tampil pas print/PDF, disembunyikan di layar biasa */}
          <div className="hidden print:block px-6 py-6">
            <h1 className="text-xl font-bold text-teal-900 mb-1">Rekap SOP — Portal SOP Kabupaten Indragiri Hulu</h1>
            <p className="text-xs text-gray-500 mb-1">
              Dicetak: {new Date().toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" })}
            </p>
            <p className="text-xs text-gray-500 mb-4">
              Filter: {opdFilter || "Semua OPD"}
              {statusFilter ? `, Status: ${statusFilter}` : ""}
              {onlyPerluDitinjau ? ", Hanya perlu ditinjau" : ""}
              {" — "}
              {results.length} SOP
            </p>
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="border-b-2 border-teal-900">
                  <th className="text-left py-1.5 pr-2">Nomor</th>
                  <th className="text-left py-1.5 pr-2">Judul</th>
                  <th className="text-left py-1.5 pr-2">OPD</th>
                  <th className="text-left py-1.5 pr-2">Efektif</th>
                  <th className="text-left py-1.5 pr-2">Status</th>
                </tr>
              </thead>
              <tbody>
                {results.map((s) => (
                  <tr key={s.id} className="border-b border-stone-200">
                    <td className="py-1.5 pr-2 font-mono">{s.nomor}</td>
                    <td className="py-1.5 pr-2">{s.judul}</td>
                    <td className="py-1.5 pr-2">{s.opd}</td>
                    <td className="py-1.5 pr-2">{formatTanggal(s.tglEfektif)}</td>
                    <td className="py-1.5 pr-2">{s.status}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      <footer className="border-t border-stone-200 mt-8">
        <div className="max-w-4xl mx-auto px-6 py-6 text-xs text-stone-400 flex flex-col sm:flex-row items-center justify-between gap-2">
          <span>Bagian Organisasi Sekretariat Daerah Kabupaten Indragiri Hulu</span>
          <div className="flex items-center gap-3">
            <span>Data disinkronkan berkala dari pengajuan OPD</span>
            <button
              onClick={() => setShowVerifikatorLogin(true)}
              className="text-stone-400 hover:text-teal-800 underline transition-colors"
            >
              Login Verifikator
            </button>
          </div>
        </div>
        <div className="max-w-4xl mx-auto px-6 pb-4 text-xs text-stone-300">
          <a
            href="https://freevectormaps.com/indonesia/ID-EPS-02-7001?ref=atr"
            target="_blank"
            rel="noreferrer"
            className="hover:text-teal-800 hover:underline transition-colors"
          >
            Peta latar oleh FreeVectorMaps.com
          </a>
        </div>
      </footer>

      {showLogin && (
        <LoginModal opdList={opdList} onClose={() => setShowLogin(false)} onLogin={handleLogin} />
      )}

      {showVerifikatorLogin && (
        <VerifikatorLoginModal
          onClose={() => setShowVerifikatorLogin(false)}
          onLogin={handleVerifikatorLogin}
        />
      )}
    </div>
  );
}
