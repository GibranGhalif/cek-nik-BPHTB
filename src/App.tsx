import { useState, useRef } from "react";
import { parseNIK, searchByName, getStatistics, NIKInfo, parseExcelBuffer, downloadExcelTemplate, exportDatabaseToExcel, getBulanName } from "./utils/nikParser";
import { PendudukData, initialDatabasePenduduk, jenisPerolehanOptions } from "./data/wilayah";

// Icons as components
function SearchIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
    </svg>
  );
}

function UserIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
    </svg>
  );
}

function IdCardIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <rect x="2" y="5" width="20" height="14" rx="2" />
      <path strokeLinecap="round" strokeLinejoin="round" d="M2 10h20" />
      <circle cx="8" cy="15" r="2" />
      <path strokeLinecap="round" d="M14 13h4M14 16h3" />
    </svg>
  );
}

function ChartIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
    </svg>
  );
}

function DatabaseIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <ellipse cx="12" cy="5" rx="9" ry="3" />
      <path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3" />
      <path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5" />
    </svg>
  );
}

function CheckCircleIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
  );
}

function XCircleIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M10 14l2-2m0 0l2-2m-2 2l-2-2m2 2l2 2m7-2a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
  );
}

function AlertIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4.5c-.77-.833-2.694-.833-3.464 0L3.34 16.5c-.77.833.192 2.5 1.732 2.5z" />
    </svg>
  );
}

function UploadIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
    </svg>
  );
}

function DownloadIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
    </svg>
  );
}

function PlusIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4" />
    </svg>
  );
}

function TrashIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
    </svg>
  );
}

type Tab = "cek-nik" | "cari-nama" | "database" | "statistik";

export default function App() {
  const [activeTab, setActiveTab] = useState<Tab>("cek-nik");
  const [nikInput, setNikInput] = useState("");
  const [namaInput, setNamaInput] = useState("");
  const [nikResult, setNikResult] = useState<NIKInfo | null>(null);
  const [namaResults, setNamaResults] = useState<PendudukData[]>([]);
  const [selectedPenduduk, setSelectedPenduduk] = useState<PendudukData | null>(null);
  const [showModal, setShowModal] = useState(false);
  const [searched, setSearched] = useState(false);
  const [namaSearched, setNamaSearched] = useState(false);
  const [database, setDatabase] = useState<PendudukData[]>(initialDatabasePenduduk);
  const [showImportModal, setShowImportModal] = useState(false);
  const [showAddModal, setShowAddModal] = useState(false);

  const handleCekNIK = () => {
    if (!nikInput.trim()) return;
    const result = parseNIK(nikInput, database);
    setNikResult(result);
    setSearched(true);
  };

  const handleCariNama = () => {
    if (!namaInput.trim()) return;
    const results = searchByName(namaInput, database);
    setNamaResults(results);
    setNamaSearched(true);
  };

  const handleViewDetail = (penduduk: PendudukData) => {
    setSelectedPenduduk(penduduk);
    setShowModal(true);
  };

  const handleImportData = (newData: PendudukData[]) => {
    setDatabase((prev) => [...prev, ...newData]);
    setShowImportModal(false);
  };

  const handleAddData = (newData: PendudukData) => {
    setDatabase((prev) => [...prev, { ...newData, id: String(Date.now()) }]);
    setShowAddModal(false);
  };

  const handleDeleteData = (id: string) => {
    if (confirm("Apakah Anda yakin ingin menghapus data ini?")) {
      setDatabase((prev) => prev.filter((p) => p.id !== id));
    }
  };

  const handleClearDatabase = () => {
    if (confirm("Apakah Anda yakin ingin menghapus semua data? Tindakan ini tidak dapat dibatalkan.")) {
      setDatabase([]);
    }
  };

  const stats = getStatistics(database);

  const tabs: { id: Tab; label: string; icon: React.ReactNode }[] = [
    { id: "cek-nik", label: "Cek NIK", icon: <IdCardIcon /> },
    { id: "cari-nama", label: "Cari Nama", icon: <SearchIcon /> },
    { id: "database", label: "Database", icon: <DatabaseIcon /> },
    { id: "statistik", label: "Statistik", icon: <ChartIcon /> },
  ];

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-950 to-slate-900">
      {/* Header */}
      <header className="border-b border-white/10 bg-white/5 backdrop-blur-xl">
        <div className="mx-auto max-w-7xl px-4 py-4 sm:px-6">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-to-br from-blue-500 to-cyan-400 shadow-lg shadow-blue-500/25">
              <IdCardIcon className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-white">Cek NIK Indonesia</h1>
              <p className="text-xs text-blue-300/70">Sistem Validasi Nomor Induk Kependudukan</p>
            </div>
          </div>
        </div>
      </header>

      {/* Navigation Tabs */}
      <nav className="border-b border-white/10 bg-white/5 backdrop-blur-sm">
        <div className="mx-auto max-w-7xl px-4 sm:px-6">
          <div className="flex gap-1 overflow-x-auto py-2">
            {tabs.map((tab) => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`flex items-center gap-2 whitespace-nowrap rounded-lg px-4 py-2.5 text-sm font-medium transition-all ${
                  activeTab === tab.id
                    ? "bg-blue-500/20 text-blue-300 shadow-lg shadow-blue-500/10"
                    : "text-slate-400 hover:bg-white/5 hover:text-slate-300"
                }`}
              >
                {tab.icon}
                {tab.label}
              </button>
            ))}
          </div>
        </div>
      </nav>

      {/* Main Content */}
      <main className="mx-auto max-w-7xl px-4 py-6 sm:px-6 sm:py-8">
        {activeTab === "cek-nik" && (
          <CekNIKTab
            nikInput={nikInput}
            setNikInput={setNikInput}
            nikResult={nikResult}
            searched={searched}
            onCek={handleCekNIK}
            onViewDetail={handleViewDetail}
            database={database}
          />
        )}

        {activeTab === "cari-nama" && (
          <CariNamaTab
            namaInput={namaInput}
            setNamaInput={setNamaInput}
            namaResults={namaResults}
            namaSearched={namaSearched}
            onCari={handleCariNama}
            onViewDetail={handleViewDetail}
          />
        )}

        {activeTab === "database" && (
          <DatabaseTab
            penduduk={database}
            onViewDetail={handleViewDetail}
            onImport={() => setShowImportModal(true)}
            onAdd={() => setShowAddModal(true)}
            onDelete={handleDeleteData}
            onClear={handleClearDatabase}
          />
        )}

        {activeTab === "statistik" && <StatistikTab stats={stats} />}
      </main>

      {/* Detail Modal */}
      {showModal && selectedPenduduk && (
        <DetailModal penduduk={selectedPenduduk} onClose={() => setShowModal(false)} />
      )}

      {/* Import Modal */}
      {showImportModal && (
        <ImportModal onClose={() => setShowImportModal(false)} onImport={handleImportData} />
      )}

      {/* Add Modal */}
      {showAddModal && (
        <AddDataModal onClose={() => setShowAddModal(false)} onAdd={handleAddData} />
      )}

      {/* Footer */}
      <footer className="border-t border-white/10 bg-white/5 backdrop-blur-sm">
        <div className="mx-auto max-w-7xl px-4 py-4 sm:px-6">
          <p className="text-center text-xs text-slate-500">
            ⚠️ Aplikasi ini menggunakan data simulasi untuk keperluan demonstrasi. Data tidak terhubung dengan database Dukcapil.
          </p>
        </div>
      </footer>
    </div>
  );
}

// === CEK NIK TAB ===
function CekNIKTab({
  nikInput,
  setNikInput,
  nikResult,
  searched,
  onCek,
  onViewDetail,
  database,
}: {
  nikInput: string;
  setNikInput: (v: string) => void;
  nikResult: NIKInfo | null;
  searched: boolean;
  onCek: () => void;
  onViewDetail: (p: PendudukData) => void;
  database: PendudukData[];
}) {
  const sampleNIKs = database.slice(0, 4).map((p) => p.nik);

  return (
    <div className="space-y-6">
      {/* Input Card */}
      <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5 backdrop-blur-xl">
        <div className="border-b border-white/10 bg-gradient-to-r from-blue-500/10 to-cyan-500/10 px-6 py-4">
          <h2 className="flex items-center gap-2 text-lg font-semibold text-white">
            <IdCardIcon className="w-5 h-5 text-blue-400" />
            Validasi NIK
          </h2>
          <p className="mt-1 text-sm text-slate-400">Masukkan 16 digit Nomor Induk Kependudukan untuk validasi</p>
        </div>
        <div className="p-6">
          <div className="flex flex-col gap-3 sm:flex-row">
            <div className="relative flex-1">
              <input
                type="text"
                value={nikInput}
                onChange={(e) => {
                  const val = e.target.value.replace(/\D/g, "").slice(0, 16);
                  setNikInput(val);
                }}
                onKeyDown={(e) => e.key === "Enter" && onCek()}
                placeholder="Contoh: 3171021505900003"
                className="w-full rounded-xl border border-white/10 bg-white/5 px-4 py-3.5 text-white placeholder-slate-500 outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2 font-mono text-lg tracking-wider"
                maxLength={16}
              />
              <span className="absolute right-3 top-1/2 -translate-y-1/2 rounded-md bg-white/10 px-2 py-0.5 text-xs font-mono text-slate-400">
                {nikInput.length}/16
              </span>
            </div>
            <button
              onClick={onCek}
              disabled={nikInput.length === 0}
              className="flex items-center justify-center gap-2 rounded-xl bg-gradient-to-r from-blue-500 to-cyan-500 px-6 py-3.5 font-semibold text-white shadow-lg shadow-blue-500/25 transition-all hover:shadow-blue-500/40 hover:brightness-110 disabled:opacity-50 disabled:cursor-not-allowed active:scale-95"
            >
              <SearchIcon className="w-5 h-5" />
              Cek NIK
            </button>
          </div>

          {/* Quick NIK samples */}
          {sampleNIKs.length > 0 && (
            <div className="mt-4">
              <p className="text-xs text-slate-500 mb-2">Contoh NIK dari database:</p>
              <div className="flex flex-wrap gap-2">
                {sampleNIKs.map((nik) => (
                  <button
                    key={nik}
                    onClick={() => setNikInput(nik)}
                    className="rounded-lg bg-white/5 border border-white/10 px-3 py-1.5 font-mono text-xs text-blue-300 hover:bg-blue-500/10 hover:border-blue-500/30 transition-all"
                  >
                    {nik}
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Results */}
      {searched && nikResult && (
        <div className="space-y-4 animate-in fade-in">
          {/* Status Card */}
          <div
            className={`overflow-hidden rounded-2xl border ${
              nikResult.valid ? "border-emerald-500/30 bg-emerald-500/5" : "border-red-500/30 bg-red-500/5"
            }`}
          >
            <div className="flex items-center gap-3 px-6 py-4">
              {nikResult.valid ? (
                <CheckCircleIcon className="w-8 h-8 text-emerald-400" />
              ) : (
                <XCircleIcon className="w-8 h-8 text-red-400" />
              )}
              <div>
                <h3 className={`text-lg font-bold ${nikResult.valid ? "text-emerald-300" : "text-red-300"}`}>
                  {nikResult.valid ? "NIK Valid" : "NIK Tidak Valid"}
                </h3>
                <p className="text-sm text-slate-400 font-mono">{nikResult.nik}</p>
              </div>
              {nikResult.penduduk && (
                <span className="ml-auto rounded-full bg-emerald-500/20 px-3 py-1 text-xs font-semibold text-emerald-300 border border-emerald-500/30">
                  ✓ Terdaftar di Database
                </span>
              )}
            </div>

            {/* Errors */}
            {nikResult.errors.length > 0 && (
              <div className="border-t border-red-500/20 px-6 py-3">
                {nikResult.errors.map((err, i) => (
                  <div key={i} className="flex items-center gap-2 py-1">
                    <AlertIcon className="w-4 h-4 text-red-400 shrink-0" />
                    <span className="text-sm text-red-300">{err}</span>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Informasi dari NIK */}
          <div className={`overflow-hidden rounded-2xl border backdrop-blur-xl ${
            nikResult.penduduk
              ? "border-emerald-500/20 bg-emerald-500/5"
              : "border-white/10 bg-white/5"
          }`}>
            <div className={`flex items-center justify-between border-b px-6 py-4 ${
              nikResult.penduduk
                ? "border-emerald-500/20 bg-gradient-to-r from-emerald-500/10 to-teal-500/10"
                : "border-white/10"
            }`}>
              <div>
                <h3 className={`font-semibold ${nikResult.penduduk ? "text-emerald-300" : "text-white"}`}>
                  Informasi dari NIK
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">
                  {nikResult.penduduk
                    ? "Data ditemukan di database"
                    : "Hasil parsing struktur NIK — data tidak ditemukan di database"}
                </p>
              </div>
              {nikResult.penduduk && (
                <button
                  onClick={() => onViewDetail(nikResult.penduduk!)}
                  className="rounded-lg bg-emerald-500/20 px-4 py-2 text-sm font-medium text-emerald-300 hover:bg-emerald-500/30 transition-all border border-emerald-500/30"
                >
                  Lihat Detail →
                </button>
              )}
            </div>

            {nikResult.penduduk ? (
              /* ── Data dari Database ── */
              <div>
                {/* Row utama: tabel sesuai kolom database */}
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-emerald-500/10 bg-white/5">
                        <th className="px-4 py-2.5 text-left font-medium text-slate-400 whitespace-nowrap">NIK</th>
                        <th className="px-4 py-2.5 text-left font-medium text-slate-400 whitespace-nowrap">Keterangan</th>
                        <th className="px-4 py-2.5 text-left font-medium text-slate-400 whitespace-nowrap">Nama</th>
                        <th className="px-4 py-2.5 text-left font-medium text-slate-400 whitespace-nowrap">Jenis Perolehan</th>
                        <th className="px-4 py-2.5 text-left font-medium text-slate-400 whitespace-nowrap">NPOTKP</th>
                        <th className="px-4 py-2.5 text-left font-medium text-slate-400 whitespace-nowrap">Tahun</th>
                        <th className="px-4 py-2.5 text-left font-medium text-slate-400 whitespace-nowrap">Bulan</th>
                        <th className="px-4 py-2.5 text-left font-medium text-slate-400 whitespace-nowrap">Tanggal</th>
                        <th className="px-4 py-2.5 text-left font-medium text-slate-400 whitespace-nowrap">Nama Kecamatan</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr className="hover:bg-white/5">
                        <td className="px-4 py-3 font-mono text-blue-300 text-xs whitespace-nowrap">{nikResult.penduduk.nik}</td>
                        <td className="px-4 py-3 whitespace-nowrap">
                          {nikResult.penduduk.keterangan ? (
                            <span className="inline-flex items-center rounded-md bg-cyan-500/20 px-2 py-0.5 text-xs text-cyan-300 border border-cyan-500/30">
                              {nikResult.penduduk.keterangan}
                            </span>
                          ) : (
                            <span className="text-slate-500 text-xs">-</span>
                          )}
                        </td>
                        <td className="px-4 py-3 text-white font-medium whitespace-nowrap">{nikResult.penduduk.nama}</td>
                        <td className="px-4 py-3 whitespace-nowrap">
                          <span className="inline-flex items-center rounded-md bg-violet-500/20 px-2 py-0.5 text-xs text-violet-300 border border-violet-500/30">
                            {nikResult.penduduk.jenisPerolehan}
                          </span>
                        </td>
                        <td className="px-4 py-3 font-mono text-xs text-slate-300 whitespace-nowrap">{nikResult.penduduk.npotkp}</td>
                        <td className="px-4 py-3 text-slate-300 whitespace-nowrap">{nikResult.penduduk.tahun}</td>
                        <td className="px-4 py-3 text-slate-300 whitespace-nowrap">{getBulanName(nikResult.penduduk.bulan)}</td>
                        <td className="px-4 py-3 text-slate-300 whitespace-nowrap">{nikResult.penduduk.tanggal}</td>
                        <td className="px-4 py-3 text-slate-300 whitespace-nowrap">{nikResult.penduduk.namaKecamatan}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* Info tambahan parsing NIK */}
                <div className="border-t border-emerald-500/10">
                  <div className="px-6 py-3 bg-white/[0.02]">
                    <p className="text-[11px] text-slate-500 font-medium uppercase tracking-wider mb-2">Info Parsing NIK</p>
                  </div>
                  <div className="grid grid-cols-2 gap-px bg-white/5 sm:grid-cols-3 lg:grid-cols-4">
                    <InfoCell label="Provinsi" value={nikResult.provinsi} code={nikResult.kodeProvinsi} icon="🏛️" />
                    <InfoCell label="Kabupaten/Kota" value={nikResult.kabupaten} code={nikResult.kodeKabupaten} icon="🏙️" />
                    <InfoCell label="Tanggal Lahir" value={nikResult.tanggalLahir} icon="📅" />
                    <InfoCell
                      label="Jenis Kelamin"
                      value={nikResult.jenisKelamin}
                      icon={nikResult.jenisKelamin === "Laki-laki" ? "👨" : "👩"}
                    />
                  </div>
                </div>
              </div>
            ) : (
              /* ── NIK tidak ada di database: tampilkan parsing saja ── */
              <div className="grid grid-cols-1 gap-px bg-white/5 sm:grid-cols-2 lg:grid-cols-3">
                <InfoCell label="Provinsi" value={nikResult.provinsi} code={nikResult.kodeProvinsi} icon="🏛️" />
                <InfoCell label="Kabupaten/Kota" value={nikResult.kabupaten} code={nikResult.kodeKabupaten} icon="🏙️" />
                <InfoCell label="Kecamatan" value={nikResult.kecamatan} code={nikResult.kodeKecamatan} icon="📍" />
                <InfoCell label="Tanggal Lahir" value={nikResult.tanggalLahir} icon="📅" />
                <InfoCell
                  label="Jenis Kelamin"
                  value={nikResult.jenisKelamin}
                  icon={nikResult.jenisKelamin === "Laki-laki" ? "👨" : "👩"}
                />
                <InfoCell label="Usia" value={nikResult.usia > 0 ? `${nikResult.usia} tahun` : "-"} icon="🎂" />
                <InfoCell label="Nomor Urut" value={nikResult.urutanRegistrasi} icon="🔢" />
              </div>
            )}
          </div>

          {/* NIK Structure Explanation */}
          <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5 backdrop-blur-xl">
            <div className="border-b border-white/10 px-6 py-4">
              <h3 className="font-semibold text-white">Struktur NIK</h3>
              <p className="text-xs text-slate-400 mt-0.5">Penjelasan setiap bagian dari NIK</p>
            </div>
            <div className="p-6">
              <div className="flex flex-wrap justify-center gap-1 font-mono text-lg">
                {nikResult.nik.length === 16 ? (
                  <>
                    <NikSegment chars={nikResult.nik.substring(0, 2)} label="Provinsi" color="from-blue-500 to-blue-600" />
                    <NikSegment chars={nikResult.nik.substring(2, 4)} label="Kab/Kota" color="from-violet-500 to-violet-600" />
                    <NikSegment chars={nikResult.nik.substring(4, 6)} label="Kecamatan" color="from-purple-500 to-purple-600" />
                    <NikSegment chars={nikResult.nik.substring(6, 8)} label="Tgl Lahir" color="from-amber-500 to-amber-600" />
                    <NikSegment chars={nikResult.nik.substring(8, 10)} label="Bln Lahir" color="from-orange-500 to-orange-600" />
                    <NikSegment chars={nikResult.nik.substring(10, 12)} label="Thn Lahir" color="from-rose-500 to-rose-600" />
                    <NikSegment chars={nikResult.nik.substring(12, 16)} label="No. Urut" color="from-emerald-500 to-emerald-600" />
                  </>
                ) : (
                  <span className="text-slate-400 text-sm">NIK harus 16 digit untuk melihat struktur</span>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function NikSegment({ chars, label, color }: { chars: string; label: string; color: string }) {
  return (
    <div className="flex flex-col items-center gap-1.5">
      <div className={`rounded-lg bg-gradient-to-b ${color} px-2.5 py-2 text-white font-bold tracking-widest shadow-lg`}>
        {chars}
      </div>
      <span className="text-[10px] font-medium text-slate-400">{label}</span>
    </div>
  );
}

function InfoCell({ label, value, code, icon }: { label: string; value: string; code?: string; icon?: string }) {
  return (
    <div className="bg-white/5 px-6 py-4">
      <div className="flex items-center gap-2 text-xs text-slate-400 mb-1">
        {icon && <span>{icon}</span>}
        {label}
        {code && <span className="ml-auto font-mono text-blue-400/60">{code}</span>}
      </div>
      <p className="font-medium text-white">{value}</p>
    </div>
  );
}

// === CARI NAMA TAB ===
function CariNamaTab({
  namaInput,
  setNamaInput,
  namaResults,
  namaSearched,
  onCari,
  onViewDetail,
}: {
  namaInput: string;
  setNamaInput: (v: string) => void;
  namaResults: PendudukData[];
  namaSearched: boolean;
  onCari: () => void;
  onViewDetail: (p: PendudukData) => void;
}) {
  return (
    <div className="space-y-6">
      <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5 backdrop-blur-xl">
        <div className="border-b border-white/10 bg-gradient-to-r from-violet-500/10 to-purple-500/10 px-6 py-4">
          <h2 className="flex items-center gap-2 text-lg font-semibold text-white">
            <UserIcon className="w-5 h-5 text-violet-400" />
            Cari Berdasarkan Nama
          </h2>
          <p className="mt-1 text-sm text-slate-400">Cari data penduduk berdasarkan nama lengkap</p>
        </div>
        <div className="p-6">
          <div className="flex flex-col gap-3 sm:flex-row">
            <input
              type="text"
              value={namaInput}
              onChange={(e) => setNamaInput(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && onCari()}
              placeholder="Masukkan nama penduduk..."
              className="flex-1 rounded-xl border border-white/10 bg-white/5 px-4 py-3.5 text-white placeholder-slate-500 outline-none ring-violet-500/50 transition-all focus:border-violet-500/50 focus:ring-2"
            />
            <button
              onClick={onCari}
              disabled={namaInput.trim().length === 0}
              className="flex items-center justify-center gap-2 rounded-xl bg-gradient-to-r from-violet-500 to-purple-500 px-6 py-3.5 font-semibold text-white shadow-lg shadow-violet-500/25 transition-all hover:shadow-violet-500/40 hover:brightness-110 disabled:opacity-50 disabled:cursor-not-allowed active:scale-95"
            >
              <SearchIcon className="w-5 h-5" />
              Cari
            </button>
          </div>
        </div>
      </div>

      {namaSearched && (
        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-medium text-slate-400">
              Hasil Pencarian: <span className="text-white">{namaResults.length}</span> data ditemukan
            </h3>
          </div>

          {namaResults.length === 0 ? (
            <div className="rounded-2xl border border-white/10 bg-white/5 p-12 text-center">
              <div className="text-4xl mb-3">🔍</div>
              <p className="text-slate-400">Tidak ada data ditemukan untuk "{namaInput}"</p>
            </div>
          ) : (
            <div className="space-y-3">
              {namaResults.map((p) => (
                <PendudukCard key={p.id} penduduk={p} onViewDetail={onViewDetail} />
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// === DATABASE TAB ===
function DatabaseTab({
  penduduk,
  onViewDetail,
  onImport,
  onAdd,
  onDelete,
  onClear,
}: {
  penduduk: PendudukData[];
  onViewDetail: (p: PendudukData) => void;
  onImport: () => void;
  onAdd: () => void;
  onDelete: (id: string) => void;
  onClear: () => void;
}) {
  const [filter, setFilter] = useState("");

  const filtered = filter
    ? penduduk.filter(
        (p) =>
          p.nama.toLowerCase().includes(filter.toLowerCase()) ||
          p.nik.includes(filter) ||
          p.namaKecamatan.toLowerCase().includes(filter.toLowerCase()) ||
          p.jenisPerolehan.toLowerCase().includes(filter.toLowerCase())
      )
    : penduduk;

  // Count NIK occurrences
  const nikCounts: Record<string, number> = {};
  penduduk.forEach((p) => {
    nikCounts[p.nik] = (nikCounts[p.nik] || 0) + 1;
  });

  return (
    <div className="space-y-6">
      <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5 backdrop-blur-xl">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-white/10 bg-gradient-to-r from-cyan-500/10 to-teal-500/10 px-6 py-4">
          <div>
            <h2 className="flex items-center gap-2 text-lg font-semibold text-white">
              <DatabaseIcon className="w-5 h-5 text-cyan-400" />
              Database Penduduk
            </h2>
            <p className="mt-1 text-sm text-slate-400">Total {penduduk.length} data penduduk</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button
              onClick={onImport}
              className="flex items-center gap-2 rounded-lg bg-emerald-500/20 px-4 py-2 text-sm font-medium text-emerald-300 hover:bg-emerald-500/30 transition-all border border-emerald-500/30"
            >
              <UploadIcon className="w-4 h-4" />
              Import Excel
            </button>
            {penduduk.length > 0 && (
              <button
                onClick={() => exportDatabaseToExcel(penduduk)}
                className="flex items-center gap-2 rounded-lg bg-cyan-500/20 px-4 py-2 text-sm font-medium text-cyan-300 hover:bg-cyan-500/30 transition-all border border-cyan-500/30"
              >
                <DownloadIcon className="w-4 h-4" />
                Export Excel
              </button>
            )}
            <button
              onClick={onAdd}
              className="flex items-center gap-2 rounded-lg bg-blue-500/20 px-4 py-2 text-sm font-medium text-blue-300 hover:bg-blue-500/30 transition-all border border-blue-500/30"
            >
              <PlusIcon className="w-4 h-4" />
              Tambah Data
            </button>
            {penduduk.length > 0 && (
              <button
                onClick={onClear}
                className="flex items-center gap-2 rounded-lg bg-red-500/20 px-4 py-2 text-sm font-medium text-red-300 hover:bg-red-500/30 transition-all border border-red-500/30"
              >
                <TrashIcon className="w-4 h-4" />
                Hapus Semua
              </button>
            )}
          </div>
        </div>
        <div className="p-4">
          <input
            type="text"
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
            placeholder="Filter berdasarkan nama, NIK, kecamatan, atau jenis perolehan..."
            className="w-full rounded-xl border border-white/10 bg-white/5 px-4 py-3 text-white placeholder-slate-500 outline-none ring-cyan-500/50 transition-all focus:border-cyan-500/50 focus:ring-2 text-sm"
          />
        </div>
      </div>

      {/* Table */}
      {penduduk.length > 0 ? (
        <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5 backdrop-blur-xl">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-white/10 bg-white/5">
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">No</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">NIK</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">Jml NIK</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">Keterangan</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">Nama</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">Jenis Perolehan</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">NPOTKP</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">Tahun</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">Bulan</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">Tanggal</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">Nama Kecamatan</th>
                  <th className="px-3 py-3 text-left font-medium text-slate-400 whitespace-nowrap">Aksi</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((p, i) => (
                  <tr key={p.id} className="border-b border-white/5 hover:bg-white/5 transition-colors">
                    <td className="px-3 py-3 text-slate-400">{i + 1}</td>
                    <td className="px-3 py-3 font-mono text-blue-300 text-xs">{p.nik}</td>
                    <td className="px-3 py-3">
                      <span
                        className={`inline-flex items-center justify-center min-w-[24px] rounded-full px-2 py-0.5 text-xs font-medium ${
                          nikCounts[p.nik] > 1
                            ? "bg-amber-500/20 text-amber-300 border border-amber-500/30"
                            : "bg-emerald-500/20 text-emerald-300 border border-emerald-500/30"
                        }`}
                      >
                        {nikCounts[p.nik]}
                      </span>
                    </td>
                    <td className="px-3 py-3">
                      {p.keterangan ? (
                        <span className="inline-flex items-center rounded-md bg-cyan-500/20 px-2 py-0.5 text-xs text-cyan-300 border border-cyan-500/30">
                          {p.keterangan}
                        </span>
                      ) : (
                        <span className="text-slate-500 text-xs">-</span>
                      )}
                    </td>
                    <td className="px-3 py-3 text-white font-medium">{p.nama}</td>
                    <td className="px-3 py-3">
                      <span className="inline-flex items-center rounded-md bg-violet-500/20 px-2 py-0.5 text-xs text-violet-300 border border-violet-500/30">
                        {p.jenisPerolehan}
                      </span>
                    </td>
                    <td className="px-3 py-3 font-mono text-xs text-slate-300">{p.npotkp}</td>
                    <td className="px-3 py-3 text-slate-300">{p.tahun}</td>
                    <td className="px-3 py-3 text-slate-300">{getBulanName(p.bulan)}</td>
                    <td className="px-3 py-3 text-slate-300">{p.tanggal}</td>
                    <td className="px-3 py-3 text-slate-300">{p.namaKecamatan}</td>
                    <td className="px-3 py-3">
                      <div className="flex gap-1">
                        <button
                          onClick={() => onViewDetail(p)}
                          className="rounded-lg bg-blue-500/20 px-2 py-1.5 text-xs font-medium text-blue-300 hover:bg-blue-500/30 transition-all border border-blue-500/30"
                        >
                          Detail
                        </button>
                        <button
                          onClick={() => onDelete(p.id)}
                          className="rounded-lg bg-red-500/20 px-2 py-1.5 text-xs font-medium text-red-300 hover:bg-red-500/30 transition-all border border-red-500/30"
                        >
                          <TrashIcon className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div className="rounded-2xl border border-white/10 bg-white/5 p-12 text-center">
          <div className="text-4xl mb-3">📭</div>
          <p className="text-slate-400 mb-4">Database kosong. Silakan tambah data atau import dari file Excel.</p>
          <div className="flex justify-center gap-3">
            <button
              onClick={onImport}
              className="flex items-center gap-2 rounded-lg bg-emerald-500/20 px-4 py-2 text-sm font-medium text-emerald-300 hover:bg-emerald-500/30 transition-all border border-emerald-500/30"
            >
              <UploadIcon className="w-4 h-4" />
              Import Excel
            </button>
            <button
              onClick={onAdd}
              className="flex items-center gap-2 rounded-lg bg-blue-500/20 px-4 py-2 text-sm font-medium text-blue-300 hover:bg-blue-500/30 transition-all border border-blue-500/30"
            >
              <PlusIcon className="w-4 h-4" />
              Tambah Data
            </button>
          </div>
        </div>
      )}

      {filtered.length === 0 && penduduk.length > 0 && (
        <div className="rounded-2xl border border-white/10 bg-white/5 p-12 text-center">
          <div className="text-4xl mb-3">🔍</div>
          <p className="text-slate-400">Tidak ada data yang sesuai dengan filter "{filter}"</p>
        </div>
      )}
    </div>
  );
}

// === STATISTIK TAB ===
function StatistikTab({ stats }: { stats: ReturnType<typeof getStatistics> }) {
  return (
    <div className="space-y-6">
      {/* Summary Cards */}
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <StatCard
          title="Total Data NIK"
          value={stats.total.toString()}
          icon="📊"
          color="from-blue-500/20 to-cyan-500/20"
          border="border-blue-500/20"
        />
        <StatCard
          title="Jenis Perolehan"
          value={Object.keys(stats.jenisPerolehanCount).length.toString()}
          icon="📋"
          color="from-violet-500/20 to-purple-500/20"
          border="border-violet-500/20"
          subtitle="kategori"
        />
        <StatCard
          title="Total Kecamatan"
          value={Object.keys(stats.kecamatanCount).length.toString()}
          icon="📍"
          color="from-amber-500/20 to-orange-500/20"
          border="border-amber-500/20"
          subtitle="kecamatan"
        />
        <StatCard
          title="Rentang Tahun"
          value={Object.keys(stats.tahunCount).length.toString()}
          icon="📅"
          color="from-emerald-500/20 to-teal-500/20"
          border="border-emerald-500/20"
          subtitle="tahun"
        />
      </div>

      {/* Charts Section */}
      <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
        {/* Jenis Perolehan Distribution */}
        <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5 backdrop-blur-xl">
          <div className="border-b border-white/10 px-6 py-4">
            <h3 className="font-semibold text-white">Distribusi Jenis Perolehan</h3>
          </div>
          <div className="p-6 space-y-3">
            {Object.entries(stats.jenisPerolehanCount)
              .sort((a, b) => b[1] - a[1])
              .map(([jenis, count]) => (
                <div key={jenis}>
                  <div className="flex items-center justify-between text-sm mb-1">
                    <span className="text-slate-300 truncate max-w-[200px]">{jenis}</span>
                    <span className="font-mono text-white">{count}</span>
                  </div>
                  <div className="h-2.5 rounded-full bg-white/10 overflow-hidden">
                    <div
                      className="h-full rounded-full bg-gradient-to-r from-violet-500 to-purple-500"
                      style={{ width: `${(count / stats.total) * 100}%` }}
                    />
                  </div>
                </div>
              ))}
            {Object.keys(stats.jenisPerolehanCount).length === 0 && (
              <p className="text-slate-400 text-center py-4">Tidak ada data</p>
            )}
          </div>
        </div>

        {/* Tahun Distribution */}
        <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5 backdrop-blur-xl">
          <div className="border-b border-white/10 px-6 py-4">
            <h3 className="font-semibold text-white">Distribusi per Tahun</h3>
          </div>
          <div className="p-6 space-y-3">
            {Object.entries(stats.tahunCount)
              .sort((a, b) => b[0].localeCompare(a[0]))
              .map(([tahun, count]) => (
                <div key={tahun}>
                  <div className="flex items-center justify-between text-sm mb-1">
                    <span className="text-slate-300">📅 {tahun}</span>
                    <span className="font-mono text-white">{count}</span>
                  </div>
                  <div className="h-2.5 rounded-full bg-white/10 overflow-hidden">
                    <div
                      className="h-full rounded-full bg-gradient-to-r from-emerald-500 to-teal-500"
                      style={{ width: `${(count / stats.total) * 100}%` }}
                    />
                  </div>
                </div>
              ))}
            {Object.keys(stats.tahunCount).length === 0 && (
              <p className="text-slate-400 text-center py-4">Tidak ada data</p>
            )}
          </div>
        </div>

        {/* Kecamatan Distribution */}
        <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5 backdrop-blur-xl">
          <div className="border-b border-white/10 px-6 py-4">
            <h3 className="font-semibold text-white">Sebaran per Kecamatan</h3>
          </div>
          <div className="p-6 space-y-2 max-h-80 overflow-y-auto">
            {Object.entries(stats.kecamatanCount)
              .sort((a, b) => b[1] - a[1])
              .map(([kecamatan, count]) => (
                <div key={kecamatan} className="flex items-center gap-3">
                  <span className="text-sm text-slate-300 w-40 truncate">{kecamatan}</span>
                  <div className="flex-1 h-2 rounded-full bg-white/10 overflow-hidden">
                    <div
                      className="h-full rounded-full bg-gradient-to-r from-cyan-500 to-blue-500"
                      style={{ width: `${(count / stats.total) * 100}%` }}
                    />
                  </div>
                  <span className="font-mono text-xs text-slate-400 w-6 text-right">{count}</span>
                </div>
              ))}
            {Object.keys(stats.kecamatanCount).length === 0 && (
              <p className="text-slate-400 text-center py-4">Tidak ada data</p>
            )}
          </div>
        </div>

        {/* Bulan Distribution */}
        <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5 backdrop-blur-xl">
          <div className="border-b border-white/10 px-6 py-4">
            <h3 className="font-semibold text-white">Distribusi per Bulan</h3>
          </div>
          <div className="p-6 space-y-2 max-h-80 overflow-y-auto">
            {Object.entries(stats.bulanCount)
              .sort((a, b) => b[1] - a[1])
              .map(([bulan, count]) => (
                <div key={bulan} className="flex items-center gap-3">
                  <span className="text-sm text-slate-300 w-28">{bulan}</span>
                  <div className="flex-1 h-2 rounded-full bg-white/10 overflow-hidden">
                    <div
                      className="h-full rounded-full bg-gradient-to-r from-amber-500 to-orange-500"
                      style={{ width: `${(count / stats.total) * 100}%` }}
                    />
                  </div>
                  <span className="font-mono text-xs text-slate-400 w-6 text-right">{count}</span>
                </div>
              ))}
            {Object.keys(stats.bulanCount).length === 0 && (
              <p className="text-slate-400 text-center py-4">Tidak ada data</p>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// === SHARED COMPONENTS ===
function PendudukCard({
  penduduk,
  onViewDetail,
}: {
  penduduk: PendudukData;
  onViewDetail: (p: PendudukData) => void;
}) {
  return (
    <div className="overflow-hidden rounded-xl border border-white/10 bg-white/5 backdrop-blur-sm hover:bg-white/10 transition-all">
      <div className="flex items-start gap-4 p-4">
        <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl text-xl bg-blue-500/20 border border-blue-500/30">
          👤
        </div>
        <div className="flex-1 min-w-0">
          <h4 className="font-semibold text-white truncate">{penduduk.nama}</h4>
          <p className="font-mono text-xs text-blue-300/80 mt-0.5">{penduduk.nik}</p>
          <div className="flex flex-wrap gap-2 mt-2">
            <span className="inline-flex items-center gap-1 rounded-md bg-white/5 px-2 py-0.5 text-xs text-slate-400 border border-white/10">
              📍 {penduduk.namaKecamatan}
            </span>
            <span className="inline-flex items-center gap-1 rounded-md bg-violet-500/10 px-2 py-0.5 text-xs text-violet-300 border border-violet-500/20">
              📋 {penduduk.jenisPerolehan}
            </span>
          </div>
        </div>
        <button
          onClick={() => onViewDetail(penduduk)}
          className="shrink-0 rounded-lg bg-blue-500/20 px-3 py-2 text-xs font-medium text-blue-300 hover:bg-blue-500/30 transition-all border border-blue-500/30"
        >
          Detail
        </button>
      </div>
    </div>
  );
}

function StatCard({
  title,
  value,
  icon,
  color,
  border,
  subtitle,
}: {
  title: string;
  value: string;
  icon: string;
  color: string;
  border: string;
  subtitle?: string;
}) {
  return (
    <div className={`overflow-hidden rounded-2xl border ${border} bg-gradient-to-br ${color} backdrop-blur-xl p-6`}>
      <div className="flex items-center justify-between">
        <div>
          <p className="text-sm text-slate-400">{title}</p>
          <p className="mt-1 text-3xl font-bold text-white">{value}</p>
          {subtitle && <p className="mt-1 text-sm text-slate-400">{subtitle}</p>}
        </div>
        <span className="text-4xl">{icon}</span>
      </div>
    </div>
  );
}

function DetailModal({ penduduk, onClose }: { penduduk: PendudukData; onClose: () => void }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="absolute inset-0 bg-black/60 backdrop-blur-sm" />
      <div
        className="relative w-full max-w-lg overflow-hidden rounded-2xl border border-white/10 bg-slate-900/95 backdrop-blur-xl shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="flex items-center justify-between border-b border-white/10 bg-gradient-to-r from-blue-500/10 to-cyan-500/10 px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-12 items-center justify-center rounded-xl text-xl bg-blue-500/20 border border-blue-500/30">
              👤
            </div>
            <div>
              <h3 className="font-bold text-white">{penduduk.nama}</h3>
              <p className="font-mono text-xs text-blue-300/80">{penduduk.nik}</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="rounded-lg p-2 text-slate-400 hover:bg-white/10 hover:text-white transition-all"
          >
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>

        {/* Content */}
        <div className="p-6">
          <div className="rounded-xl bg-gradient-to-br from-blue-900/50 to-cyan-900/50 border border-blue-500/20 p-5 space-y-4">
            <div className="grid grid-cols-[auto_8px_1fr] gap-y-2 text-sm">
              <span className="text-blue-300">NIK</span>
              <span className="text-blue-300">:</span>
              <span className="font-mono font-bold text-white">{penduduk.nik}</span>

              <span className="text-blue-300">Keterangan</span>
              <span className="text-blue-300">:</span>
              <span className="text-white">
                {penduduk.keterangan ? (
                  <span className="inline-flex items-center rounded-md bg-cyan-500/20 px-2 py-0.5 text-xs text-cyan-300 border border-cyan-500/30">
                    {penduduk.keterangan}
                  </span>
                ) : (
                  <span className="text-slate-400">-</span>
                )}
              </span>

              <span className="text-blue-300">Nama</span>
              <span className="text-blue-300">:</span>
              <span className="font-semibold text-white">{penduduk.nama}</span>

              <span className="text-blue-300">Jenis Perolehan</span>
              <span className="text-blue-300">:</span>
              <span className="text-white">{penduduk.jenisPerolehan}</span>

              <span className="text-blue-300">NPOTKP</span>
              <span className="text-blue-300">:</span>
              <span className="font-mono text-white">{penduduk.npotkp}</span>

              <span className="text-blue-300">Tahun</span>
              <span className="text-blue-300">:</span>
              <span className="text-white">{penduduk.tahun}</span>

              <span className="text-blue-300">Bulan</span>
              <span className="text-blue-300">:</span>
              <span className="text-white">{getBulanName(penduduk.bulan)}</span>

              <span className="text-blue-300">Tanggal</span>
              <span className="text-blue-300">:</span>
              <span className="text-white">{penduduk.tanggal}</span>

              <span className="text-blue-300">Nama Kecamatan</span>
              <span className="text-blue-300">:</span>
              <span className="text-white">{penduduk.namaKecamatan}</span>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="border-t border-white/10 px-6 py-4 flex justify-end">
          <button
            onClick={onClose}
            className="rounded-xl bg-white/10 px-5 py-2.5 text-sm font-medium text-white hover:bg-white/20 transition-all border border-white/10"
          >
            Tutup
          </button>
        </div>
      </div>
    </div>
  );
}

// === IMPORT MODAL ===
function ImportModal({
  onClose,
  onImport,
}: {
  onClose: () => void;
  onImport: (data: PendudukData[]) => void;
}) {
  const [preview, setPreview] = useState<PendudukData[]>([]);
  const [error, setError] = useState("");
  const [fileName, setFileName] = useState("");
  const [sheetInfo, setSheetInfo] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const processFile = (file: File) => {
    const validTypes = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
      "application/x-excel",
    ];
    const validExtensions = [".xlsx", ".xls"];
    const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();

    if (!validTypes.includes(file.type) && !validExtensions.includes(ext)) {
      setError("Format file tidak didukung. Gunakan file Excel (.xlsx atau .xls)");
      return;
    }

    setIsLoading(true);
    setFileName(file.name);
    setError("");

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer;
        const result = parseExcelBuffer(buffer);

        if (result.errors.length > 0) {
          setError(result.errors.join("\n"));
          setPreview([]);
        } else {
          setPreview(result.data);
          setSheetInfo(`Sheet: "${result.sheetName}" — ${result.totalRows} baris ditemukan`);
          setError("");
        }
      } catch {
        setError("Gagal membaca file Excel. Pastikan file tidak rusak.");
        setPreview([]);
      } finally {
        setIsLoading(false);
      }
    };
    reader.onerror = () => {
      setError("Gagal membaca file.");
      setIsLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) processFile(file);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="absolute inset-0 bg-black/60 backdrop-blur-sm" />
      <div
        className="relative w-full max-w-2xl max-h-[90vh] overflow-hidden rounded-2xl border border-white/10 bg-slate-900/95 backdrop-blur-xl shadow-2xl flex flex-col"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="flex items-center justify-between border-b border-white/10 bg-gradient-to-r from-emerald-500/10 to-teal-500/10 px-6 py-4 shrink-0">
          <div>
            <h3 className="font-bold text-white flex items-center gap-2">
              <UploadIcon className="w-5 h-5 text-emerald-400" />
              Import Data dari Excel
            </h3>
            <p className="text-xs text-slate-400 mt-0.5">Upload file Excel (.xlsx / .xls) untuk import data NIK</p>
          </div>
          <button
            onClick={onClose}
            className="rounded-lg p-2 text-slate-400 hover:bg-white/10 hover:text-white transition-all"
          >
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>

        {/* Content */}
        <div className="p-6 space-y-4 overflow-y-auto flex-1">
          {/* File Upload Drag & Drop */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="block text-sm font-medium text-slate-300">Upload File Excel</label>
              <button
                onClick={() => downloadExcelTemplate()}
                className="flex items-center gap-1.5 rounded-lg bg-blue-500/20 px-3 py-1.5 text-xs font-medium text-blue-300 hover:bg-blue-500/30 transition-all border border-blue-500/30"
              >
                <DownloadIcon className="w-3.5 h-3.5" />
                Download Template .xlsx
              </button>
            </div>
            <div
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onClick={() => fileInputRef.current?.click()}
              className={`flex flex-col items-center justify-center gap-3 rounded-xl border-2 border-dashed py-8 cursor-pointer transition-all ${
                isDragging
                  ? "border-emerald-400 bg-emerald-500/10 text-emerald-300"
                  : "border-white/20 text-slate-400 hover:border-emerald-500/50 hover:text-emerald-300 hover:bg-emerald-500/5"
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileChange}
                className="hidden"
              />
              {isLoading ? (
                <>
                  <div className="h-10 w-10 animate-spin rounded-full border-4 border-emerald-500/30 border-t-emerald-400" />
                  <p className="text-sm">Membaca file...</p>
                </>
              ) : fileName ? (
                <>
                  <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-emerald-500/20 border border-emerald-500/30">
                    <ExcelIcon className="w-6 h-6 text-emerald-400" />
                  </div>
                  <div className="text-center">
                    <p className="text-sm font-medium text-emerald-300">{fileName}</p>
                    {sheetInfo && <p className="text-xs text-slate-400 mt-0.5">{sheetInfo}</p>}
                  </div>
                  <p className="text-xs text-slate-500">Klik untuk ganti file</p>
                </>
              ) : (
                <>
                  <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-white/5 border border-white/10">
                    <UploadIcon className="w-6 h-6" />
                  </div>
                  <div className="text-center">
                    <p className="text-sm font-medium">Drag & drop file Excel di sini</p>
                    <p className="text-xs text-slate-500 mt-1">atau klik untuk memilih file</p>
                  </div>
                  <div className="flex gap-2 mt-1">
                    <span className="rounded-md bg-white/10 px-2 py-0.5 text-[10px] font-mono text-slate-400">.xlsx</span>
                    <span className="rounded-md bg-white/10 px-2 py-0.5 text-[10px] font-mono text-slate-400">.xls</span>
                  </div>
                </>
              )}
            </div>
          </div>

          {/* Format Info */}
          <div className="rounded-xl bg-blue-500/5 border border-blue-500/20 px-4 py-3">
            <p className="text-xs font-medium text-blue-300 mb-1.5">📋 Format kolom yang didukung:</p>
            <div className="grid grid-cols-3 gap-1 text-[11px] text-slate-400">
              <span className="bg-white/5 rounded px-2 py-1">NIK</span>
              <span className="bg-white/5 rounded px-2 py-1">Keterangan</span>
              <span className="bg-white/5 rounded px-2 py-1">Nama</span>
              <span className="bg-white/5 rounded px-2 py-1">Jenis Perolehan</span>
              <span className="bg-white/5 rounded px-2 py-1">NPOTKP</span>
              <span className="bg-white/5 rounded px-2 py-1">Tahun</span>
              <span className="bg-white/5 rounded px-2 py-1">Bulan</span>
              <span className="bg-white/5 rounded px-2 py-1">Tanggal</span>
              <span className="bg-white/5 rounded px-2 py-1">Nama Kecamatan</span>
              <span className="bg-white/5 rounded px-2 py-1">No / ID</span>
            </div>
            <p className="text-[10px] text-slate-500 mt-1.5">* Kolom NIK wajib ada. Kolom lain opsional. Nama kolom fleksibel.</p>
          </div>

          {/* Error */}
          {error && (
            <div className="rounded-lg bg-red-500/10 border border-red-500/30 px-4 py-3 text-sm text-red-300 flex items-start gap-2">
              <AlertIcon className="w-4 h-4 shrink-0 mt-0.5" />
              <span className="whitespace-pre-wrap">{error}</span>
            </div>
          )}

          {/* Preview */}
          {preview.length > 0 && (
            <div>
              <div className="flex items-center justify-between mb-2">
                <label className="block text-sm font-medium text-emerald-300">
                  ✅ Preview: {preview.length} data siap diimport
                </label>
              </div>
              <div className="rounded-xl border border-white/10 bg-white/5 overflow-hidden max-h-52 overflow-y-auto">
                <table className="w-full text-xs">
                  <thead className="bg-white/5 sticky top-0">
                    <tr>
                      <th className="px-3 py-2 text-left text-slate-400">No</th>
                      <th className="px-3 py-2 text-left text-slate-400">NIK</th>
                      <th className="px-3 py-2 text-left text-slate-400">Keterangan</th>
                      <th className="px-3 py-2 text-left text-slate-400">Nama</th>
                      <th className="px-3 py-2 text-left text-slate-400">Jenis Perolehan</th>
                      <th className="px-3 py-2 text-left text-slate-400">Kecamatan</th>
                    </tr>
                  </thead>
                  <tbody>
                    {preview.slice(0, 20).map((p, i) => (
                      <tr key={i} className="border-t border-white/5 hover:bg-white/5">
                        <td className="px-3 py-2 text-slate-500">{i + 1}</td>
                        <td className="px-3 py-2 font-mono text-blue-300">{p.nik}</td>
                        <td className="px-3 py-2 text-cyan-300">{p.keterangan || "-"}</td>
                        <td className="px-3 py-2 text-white">{p.nama}</td>
                        <td className="px-3 py-2 text-slate-300">{p.jenisPerolehan}</td>
                        <td className="px-3 py-2 text-slate-300">{p.namaKecamatan}</td>
                      </tr>
                    ))}
                    {preview.length > 20 && (
                      <tr className="border-t border-white/5">
                        <td colSpan={6} className="px-3 py-2 text-center text-slate-500">
                          ...dan {preview.length - 20} data lainnya
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="border-t border-white/10 px-6 py-4 flex justify-end gap-3 shrink-0">
          <button
            onClick={onClose}
            className="rounded-xl bg-white/10 px-5 py-2.5 text-sm font-medium text-white hover:bg-white/20 transition-all border border-white/10"
          >
            Batal
          </button>
          <button
            onClick={() => preview.length > 0 && onImport(preview)}
            disabled={preview.length === 0}
            className="rounded-xl bg-gradient-to-r from-emerald-500 to-teal-500 px-5 py-2.5 text-sm font-semibold text-white shadow-lg shadow-emerald-500/25 transition-all hover:brightness-110 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            Import {preview.length > 0 && `(${preview.length} data)`}
          </button>
        </div>
      </div>
    </div>
  );
}

function ExcelIcon({ className = "w-5 h-5" }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={1.5}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 17V7m0 10a2 2 0 01-2 2H5a2 2 0 01-2-2V7a2 2 0 012-2h2a2 2 0 012 2m0 10a2 2 0 002 2h2a2 2 0 002-2M9 7a2 2 0 012-2h2a2 2 0 012 2m0 10V7m0 10a2 2 0 002 2h2a2 2 0 002-2V7a2 2 0 00-2-2h-2a2 2 0 00-2 2" />
    </svg>
  );
}

// === ADD DATA MODAL ===
function AddDataModal({
  onClose,
  onAdd,
}: {
  onClose: () => void;
  onAdd: (data: PendudukData) => void;
}) {
  const [formData, setFormData] = useState<Omit<PendudukData, "id">>({
    nik: "",
    keterangan: "",
    nama: "",
    jenisPerolehan: jenisPerolehanOptions[0],
    npotkp: "",
    tahun: new Date().getFullYear().toString(),
    bulan: (new Date().getMonth() + 1).toString().padStart(2, "0"),
    tanggal: new Date().getDate().toString().padStart(2, "0"),
    namaKecamatan: "",
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.nik || !formData.nama) {
      alert("NIK dan Nama harus diisi!");
      return;
    }
    onAdd({ ...formData, id: "" });
  };

  const bulanOptions = [
    { value: "01", label: "Januari" },
    { value: "02", label: "Februari" },
    { value: "03", label: "Maret" },
    { value: "04", label: "April" },
    { value: "05", label: "Mei" },
    { value: "06", label: "Juni" },
    { value: "07", label: "Juli" },
    { value: "08", label: "Agustus" },
    { value: "09", label: "September" },
    { value: "10", label: "Oktober" },
    { value: "11", label: "November" },
    { value: "12", label: "Desember" },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="absolute inset-0 bg-black/60 backdrop-blur-sm" />
      <div
        className="relative w-full max-w-lg overflow-hidden rounded-2xl border border-white/10 bg-slate-900/95 backdrop-blur-xl shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="flex items-center justify-between border-b border-white/10 bg-gradient-to-r from-blue-500/10 to-cyan-500/10 px-6 py-4">
          <div>
            <h3 className="font-bold text-white flex items-center gap-2">
              <PlusIcon className="w-5 h-5 text-blue-400" />
              Tambah Data Baru
            </h3>
          </div>
          <button
            onClick={onClose}
            className="rounded-lg p-2 text-slate-400 hover:bg-white/10 hover:text-white transition-all"
          >
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>

        {/* Form */}
        <form onSubmit={handleSubmit} className="p-6 space-y-4">
          <div>
            <label className="block text-sm font-medium text-slate-300 mb-1.5">NIK *</label>
            <input
              type="text"
              value={formData.nik}
              onChange={(e) => setFormData({ ...formData, nik: e.target.value.replace(/\D/g, "").slice(0, 16) })}
              placeholder="16 digit NIK"
              className="w-full rounded-lg border border-white/10 bg-white/5 px-4 py-2.5 text-white placeholder-slate-500 outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2 font-mono"
              maxLength={16}
              required
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-300 mb-1.5">Keterangan</label>
            <input
              type="text"
              value={formData.keterangan}
              onChange={(e) => setFormData({ ...formData, keterangan: e.target.value })}
              placeholder="Contoh: Data Valid, Perlu Verifikasi, dll"
              className="w-full rounded-lg border border-white/10 bg-white/5 px-4 py-2.5 text-white placeholder-slate-500 outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-300 mb-1.5">Nama *</label>
            <input
              type="text"
              value={formData.nama}
              onChange={(e) => setFormData({ ...formData, nama: e.target.value })}
              placeholder="Nama lengkap"
              className="w-full rounded-lg border border-white/10 bg-white/5 px-4 py-2.5 text-white placeholder-slate-500 outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2"
              required
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-300 mb-1.5">Jenis Perolehan</label>
            <select
              value={formData.jenisPerolehan}
              onChange={(e) => setFormData({ ...formData, jenisPerolehan: e.target.value })}
              className="w-full rounded-lg border border-white/10 bg-white/5 px-4 py-2.5 text-white outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2"
            >
              {jenisPerolehanOptions.map((opt) => (
                <option key={opt} value={opt} className="bg-slate-800">
                  {opt}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-300 mb-1.5">NPOTKP</label>
            <input
              type="text"
              value={formData.npotkp}
              onChange={(e) => setFormData({ ...formData, npotkp: e.target.value })}
              placeholder="NPO/XXXX/YYYY/XXXXXX"
              className="w-full rounded-lg border border-white/10 bg-white/5 px-4 py-2.5 text-white placeholder-slate-500 outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2 font-mono"
            />
          </div>

          <div className="grid grid-cols-3 gap-3">
            <div>
              <label className="block text-sm font-medium text-slate-300 mb-1.5">Tahun</label>
              <input
                type="text"
                value={formData.tahun}
                onChange={(e) => setFormData({ ...formData, tahun: e.target.value.replace(/\D/g, "").slice(0, 4) })}
                placeholder="2024"
                className="w-full rounded-lg border border-white/10 bg-white/5 px-4 py-2.5 text-white placeholder-slate-500 outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2"
                maxLength={4}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-300 mb-1.5">Bulan</label>
              <select
                value={formData.bulan}
                onChange={(e) => setFormData({ ...formData, bulan: e.target.value })}
                className="w-full rounded-lg border border-white/10 bg-white/5 px-3 py-2.5 text-white outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2"
              >
                {bulanOptions.map((opt) => (
                  <option key={opt.value} value={opt.value} className="bg-slate-800">
                    {opt.label}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-300 mb-1.5">Tanggal</label>
              <input
                type="text"
                value={formData.tanggal}
                onChange={(e) => setFormData({ ...formData, tanggal: e.target.value.replace(/\D/g, "").slice(0, 2) })}
                placeholder="01"
                className="w-full rounded-lg border border-white/10 bg-white/5 px-4 py-2.5 text-white placeholder-slate-500 outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2"
                maxLength={2}
              />
            </div>
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-300 mb-1.5">Nama Kecamatan</label>
            <input
              type="text"
              value={formData.namaKecamatan}
              onChange={(e) => setFormData({ ...formData, namaKecamatan: e.target.value })}
              placeholder="Nama kecamatan"
              className="w-full rounded-lg border border-white/10 bg-white/5 px-4 py-2.5 text-white placeholder-slate-500 outline-none ring-blue-500/50 transition-all focus:border-blue-500/50 focus:ring-2"
            />
          </div>
        </form>

        {/* Footer */}
        <div className="border-t border-white/10 px-6 py-4 flex justify-end gap-3">
          <button
            onClick={onClose}
            className="rounded-xl bg-white/10 px-5 py-2.5 text-sm font-medium text-white hover:bg-white/20 transition-all border border-white/10"
          >
            Batal
          </button>
          <button
            onClick={handleSubmit}
            className="rounded-xl bg-gradient-to-r from-blue-500 to-cyan-500 px-5 py-2.5 text-sm font-semibold text-white shadow-lg shadow-blue-500/25 transition-all hover:brightness-110"
          >
            Simpan
          </button>
        </div>
      </div>
    </div>
  );
}
