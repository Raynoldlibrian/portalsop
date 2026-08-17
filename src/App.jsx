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
} from "lucide-react";

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
      {/* tab strip — catalog-card signature */}
      <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-amber-500" />
      <div className="pl-6 pr-5 py-5">
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
            <h3 className="text-lg leading-snug font-semibold text-teal-900 mb-3" style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}>
              {sop.judul}
            </h3>
            <dl className="grid grid-cols-3 gap-x-4 gap-y-1 text-xs">
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
            <a
              href="#"
              onClick={(e) => {
                e.preventDefault();
                onOpen();
              }}
              className="inline-flex items-center gap-1.5 rounded border border-teal-800 text-teal-900 text-xs font-semibold px-3 py-2 hover:bg-teal-700 hover:text-white transition-colors"
            >
              <Download size={13} strokeWidth={2.3} />
              Lihat
            </a>
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

function SopDetailPage({ sop, isOperator, onBack, onEdit, onRevisi }) {
  const berlaku = sop.status === "Berlaku";
  return (
    <div className="max-w-4xl mx-auto px-6 py-8">
      <button
        onClick={onBack}
        className="inline-flex items-center gap-1.5 text-sm font-medium text-gray-500 hover:text-teal-800 transition-colors mb-5"
      >
        <ArrowLeft size={15} strokeWidth={2.2} />
        Kembali ke pencarian
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
          <a
            href="#"
            onClick={(e) => e.preventDefault()}
            className="inline-flex items-center gap-1.5 text-sm font-semibold text-white bg-teal-800 rounded-md px-4 py-2.5 hover:bg-teal-700 transition-colors"
          >
            <Eye size={14} strokeWidth={2.2} />
            Lihat PDF
          </a>
          <a
            href="#"
            onClick={(e) => e.preventDefault()}
            className="inline-flex items-center gap-1.5 text-sm font-semibold text-teal-900 border border-teal-800 rounded-md px-4 py-2.5 hover:bg-teal-50 transition-colors"
          >
            <Download size={14} strokeWidth={2.2} />
            Unduh
          </a>
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
    </div>
  );
}

const STAT_COLORS = {
  emerald: { bg: "bg-emerald-50", icon: "text-emerald-700", value: "text-emerald-800" },
  blue: { bg: "bg-sky-50", icon: "text-sky-700", value: "text-sky-800" },
  slate: { bg: "bg-stone-100", icon: "text-stone-600", value: "text-stone-800" },
};

function StatCard({ icon: Icon, label, value, color }) {
  const c = STAT_COLORS[color] || STAT_COLORS.blue;
  return (
    <div className="flex flex-col items-start gap-2 bg-white border border-stone-200 rounded-xl px-3 py-3 sm:px-4 sm:py-4 min-w-0">
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

function VerifikasiCard({ sop, onSelesai }) {
  const [catatan, setCatatan] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const kirim = async (statusVerifikasi) => {
    if (statusVerifikasi === "Revisi" && !catatan.trim()) {
      setError("Isi catatan dulu, biar OPD tau apa yang perlu diperbaiki");
      return;
    }
    setError("");
    setBusy(true);
    try {
      await apiPost("verifikasi", {
        id_sop: sop.id,
        status_verifikasi: statusVerifikasi,
        catatan_verifikasi: catatan.trim(),
      });
      onSelesai();
    } catch (err) {
      setError(err.message || "Gagal mengirim, coba lagi.");
      setBusy(false);
    }
  };

  return (
    <div className="bg-white border border-stone-200 rounded-xl p-5">
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
        className="text-base leading-snug font-semibold text-teal-900 mb-3"
        style={{ fontFamily: "'Source Serif 4', Georgia, serif" }}
      >
        {sop.judul}
      </h3>

      <dl className="grid grid-cols-2 sm:grid-cols-4 gap-x-4 gap-y-2 text-xs mb-4">
        <div>
          <dt className="text-stone-400">Bidang/Bagian</dt>
          <dd className="text-gray-700 font-medium">{sop.bidang}</dd>
        </div>
        <div>
          <dt className="text-stone-400">Disahkan oleh</dt>
          <dd className="text-gray-700 font-medium">{sop.disahkanOleh}</dd>
        </div>
        <div>
          <dt className="text-stone-400">Tgl Efektif</dt>
          <dd className="text-gray-700 font-medium">{formatTanggal(sop.tglEfektif)}</dd>
        </div>
        <div>
          <dt className="text-stone-400">Status</dt>
          <dd className="text-gray-700 font-medium">{sop.status}</dd>
        </div>
      </dl>

      <a
        href={sop.linkDrive}
        target="_blank"
        rel="noreferrer"
        className="inline-flex items-center gap-1.5 text-xs font-semibold text-teal-900 border border-teal-800 rounded-md px-3 py-1.5 hover:bg-teal-50 transition-colors mb-4"
      >
        <Eye size={12} strokeWidth={2.2} />
        Lihat PDF SOP
      </a>

      <div>
        <label className="block text-xs font-semibold text-gray-700 mb-1.5">
          Catatan <span className="text-stone-400 font-normal">(wajib kalau minta revisi)</span>
        </label>
        <textarea
          rows={2}
          value={catatan}
          onChange={(e) => setCatatan(e.target.value)}
          placeholder="cth. Link Drive belum bisa diakses publik…"
          className="w-full rounded-md border border-stone-200 bg-white px-3 py-2.5 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800 resize-none"
        />
      </div>

      {error && <p className="mt-2 text-xs text-red-600">{error}</p>}

      <div className="mt-3 flex items-center justify-end gap-2">
        <button
          type="button"
          disabled={busy}
          onClick={() => kirim("Revisi")}
          className="inline-flex items-center gap-1.5 text-sm font-semibold text-amber-700 border border-amber-400 rounded-md px-4 py-2 hover:bg-amber-50 transition-colors disabled:opacity-50"
        >
          <History size={14} strokeWidth={2.2} />
          Minta Revisi
        </button>
        <button
          type="button"
          disabled={busy}
          onClick={() => kirim("Terverifikasi")}
          className="inline-flex items-center gap-1.5 text-sm font-semibold text-white bg-teal-800 rounded-md px-4 py-2 hover:bg-teal-700 transition-colors disabled:opacity-50"
        >
          <CheckCircle2 size={14} strokeWidth={2.2} />
          Verifikasi
        </button>
      </div>
    </div>
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

      {sop.statusVerifikasi === "Revisi" && sop.catatanVerifikasi && (
        <div className="rounded-md border border-red-200 bg-red-50 px-3 py-2.5 mb-3">
          <p className="text-xs font-semibold text-red-700 mb-1">Catatan dari Bagian Organisasi:</p>
          <p className="text-xs text-red-700">{sop.catatanVerifikasi}</p>
        </div>
      )}

      <div className="flex items-center gap-2">
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

function VerifikasiPage({ verifikatorName, onLogout, onBack }) {
  const [pending, setPending] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

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
            <button
              onClick={onLogout}
              className="text-xs font-medium text-teal-100 bg-white bg-opacity-10 hover:bg-opacity-20 rounded-full px-3 py-1.5 transition-colors"
            >
              Keluar
            </button>
            <button
              onClick={onBack}
              className="text-xs font-medium text-teal-100 hover:text-white transition-colors"
            >
              ← Ke portal publik
            </button>
          </div>
        </div>
      </div>

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

        <div className="grid gap-4">
          {pending.map((sop) => (
            <VerifikasiCard key={sop.id} sop={sop} onSelesai={load} />
          ))}
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
    return sopList.filter((s) => {
      const matchesQuery =
        query.trim() === "" ||
        s.judul.toLowerCase().includes(query.toLowerCase()) ||
        s.nomor.toLowerCase().includes(query.toLowerCase());
      const matchesOpd = opdFilter === "" || s.opd === opdFilter;
      return matchesQuery && matchesOpd;
    });
  }, [sopList, query, opdFilter]);

  const latest = useMemo(() => {
    return [...sopList].sort((a, b) => (a.diverifikasi < b.diverifikasi ? 1 : -1)).slice(0, 5);
  }, [sopList]);

  const stats = useMemo(() => {
    const opdBerpartisipasi = new Set(sopList.map((s) => s.idOpd)).size;
    return [
      { label: "SOP Terverifikasi", value: sopList.length, icon: ClipboardList, color: "emerald" },
      { label: "OPD Berpartisipasi", value: opdBerpartisipasi, icon: Landmark, color: "blue" },
      { label: "Total OPD di Inhu", value: opdList.length, icon: Layers, color: "slate" },
    ];
  }, [sopList, opdList]);

  if (view === "verifikasi" && verifikatorName) {
    return (
      <VerifikasiPage
        verifikatorName={verifikatorName}
        onLogout={handleVerifikatorLogout}
        onBack={() => setView("list")}
      />
    );
  }

  return (
    <div className="min-h-screen bg-stone-50 font-sans" style={{ fontFamily: "'Inter', system-ui, sans-serif" }}>
      {/* Header — teal bar with logo lockup, matching Patuhdiri/MANDALA branding */}
      <div className="bg-teal-900">
        <div className="max-w-4xl mx-auto px-6 py-5 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="shrink-0 w-12 h-12 rounded-full bg-white flex items-center justify-center overflow-hidden">
              <img src="https://i.imgur.com/jk45oWc.png" alt="Logo Kabupaten Indragiri Hulu" className="w-8 h-8 object-contain" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-white leading-none">Portal SOP</h1>
              <p className="text-sm text-teal-100 mt-1">Registri Standar Operasional Prosedur</p>
              <p className="text-xs italic text-teal-300 mt-0.5">
                Bagian Organisasi Sekretariat Daerah Kabupaten Indragiri Hulu
              </p>
            </div>
          </div>
          {loggedInOpd ? (
            <div className="flex items-center gap-2 shrink-0 flex-wrap justify-end">
              <span className="inline-flex items-center gap-1.5 text-xs font-medium text-teal-900 bg-amber-400 rounded-full px-3 py-1.5 whitespace-nowrap">
                <ShieldCheck size={13} strokeWidth={2.3} />
                {loggedInOpd.nama_opd}
              </span>
              <button
                onClick={() => setView("sopsaya")}
                className="text-xs font-medium text-teal-100 bg-white bg-opacity-10 hover:bg-opacity-20 rounded-full px-3 py-1.5 whitespace-nowrap transition-colors"
              >
                SOP Saya
              </button>
              <button
                onClick={handleLogout}
                className="text-xs font-medium text-teal-100 bg-white bg-opacity-10 hover:bg-opacity-20 rounded-full px-3 py-1.5 whitespace-nowrap transition-colors"
              >
                Keluar
              </button>
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
      <header className="relative overflow-hidden border-b border-stone-200">
        <img
          src="https://i.imgur.com/NTqvgbH.png"
          alt=""
          className="absolute inset-0 w-full h-full object-cover"
        />
        <div className="absolute inset-0 bg-teal-900 bg-opacity-80" />
        <div className="relative max-w-4xl mx-auto px-6 pt-10 pb-8">
          <div className="flex items-center gap-2 text-xs font-semibold tracking-widest uppercase text-amber-400 mb-2">
            <FileText size={14} strokeWidth={2.3} />
            Registri SOP
          </div>
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

          {/* Search bar */}
          <div className="relative">
            <Search
              size={17}
              strokeWidth={2.2}
              className="absolute left-4 top-1/2 -translate-y-1/2 text-stone-400"
            />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              type="text"
              placeholder="Cari judul atau nomor SOP…"
              className="w-full pl-11 pr-4 py-3.5 rounded-md border border-stone-200 bg-stone-50 text-sm text-teal-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-2 focus:ring-teal-800 focus:ring-opacity-30 focus:border-teal-800 transition-all"
            />
          </div>

          {/* OPD filter dropdown + Ajukan SOP Baru */}
          <div className="mt-3 flex items-center justify-between gap-2 flex-wrap">
            <div className="flex items-center gap-2 flex-wrap">
              <div className="relative">
                <button
                  onClick={() => setOpdOpen((v) => !v)}
                  className="inline-flex items-center gap-2 text-sm font-medium text-gray-700 border border-stone-200 rounded-md pl-3 pr-2.5 py-2 bg-white hover:border-teal-800 transition-colors"
                >
                  <Building2 size={14} strokeWidth={2.2} className="text-stone-400" />
                  {opdFilter || "Semua OPD"}
                  <ChevronDown size={14} strokeWidth={2.2} className="text-stone-400" />
                </button>
                {opdOpen && (
                  <div className="absolute z-10 mt-1.5 w-64 bg-white border border-stone-200 rounded-md shadow-lg py-1.5 max-h-64 overflow-auto">
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
            </div>

            {loggedInOpd && (
              <button
                onClick={openTambah}
                className="inline-flex items-center gap-2 bg-teal-800 text-white font-semibold text-sm px-4 py-2 rounded-md hover:bg-teal-700 transition-colors"
              >
                <Plus size={16} />
                Ajukan SOP Baru
              </button>
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
      ) : (
        <>
          {/* Dashboard stats */}
          <section className="max-w-4xl mx-auto px-6 pt-8">
            <div className="grid grid-cols-3 gap-2 sm:gap-3">
              {stats.map((s) => (
                <StatCard key={s.label} {...s} />
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
          <main className="max-w-4xl mx-auto px-6 py-8">
            <div className="flex items-center gap-2 mb-3 pt-2 border-t border-stone-200">
              <ClipboardList size={15} strokeWidth={2.2} className="text-teal-900" />
              <h2 className="text-sm font-semibold text-teal-900">Jelajahi Semua SOP</h2>
            </div>
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2 text-sm text-gray-500">
                <CalendarCheck size={14} strokeWidth={2.2} className="text-green-700" />
                <span>
                  <strong className="text-teal-900">{results.length}</strong> SOP terverifikasi ditemukan
                </span>
              </div>
            </div>

            {results.length === 0 ? (
              <div className="text-center py-20 border border-dashed border-stone-200 rounded-md">
                <FileText size={28} strokeWidth={1.5} className="mx-auto text-teal-100 mb-3" />
                <p className="text-gray-500 text-sm mb-1">Tidak ada SOP yang cocok dengan pencarian.</p>
                <p className="text-stone-400 text-xs">Coba kata kunci lain atau ubah filter OPD.</p>
              </div>
            ) : (
              <div className="grid gap-3">
                {results.map((sop) => (
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
            )}
          </main>
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
