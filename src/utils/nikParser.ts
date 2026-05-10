import * as XLSX from "xlsx";
import { wilayahIndonesia, kecamatanDatabase, PendudukData } from "../data/wilayah";

export interface NIKInfo {
  valid: boolean;
  nik: string;
  provinsi: string;
  kabupaten: string;
  kecamatan: string;
  kodeProvinsi: string;
  kodeKabupaten: string;
  kodeKecamatan: string;
  tanggalLahir: string;
  jenisKelamin: string;
  usia: number;
  urutanRegistrasi: string;
  errors: string[];
  penduduk?: PendudukData;
}

export function parseNIK(nik: string, database: PendudukData[]): NIKInfo {
  const errors: string[] = [];
  const cleanNIK = nik.replace(/\s/g, "");

  const result: NIKInfo = {
    valid: false,
    nik: cleanNIK,
    provinsi: "-",
    kabupaten: "-",
    kecamatan: "-",
    kodeProvinsi: "-",
    kodeKabupaten: "-",
    kodeKecamatan: "-",
    tanggalLahir: "-",
    jenisKelamin: "-",
    usia: 0,
    urutanRegistrasi: "-",
    errors: [],
  };

  // Check length
  if (cleanNIK.length !== 16) {
    errors.push(`NIK harus 16 digit, Anda memasukkan ${cleanNIK.length} digit`);
  }

  // Check if all digits
  if (!/^\d+$/.test(cleanNIK)) {
    errors.push("NIK hanya boleh berisi angka");
  }

  if (errors.length > 0 && cleanNIK.length < 6) {
    result.errors = errors;
    return result;
  }

  // Parse kode wilayah
  const kodeProvinsi = cleanNIK.substring(0, 2);
  const kodeKabupaten = cleanNIK.substring(0, 4);
  const kodeKecamatan = cleanNIK.substring(0, 6);

  result.kodeProvinsi = kodeProvinsi;
  result.kodeKabupaten = kodeKabupaten;
  result.kodeKecamatan = kodeKecamatan;

  // Lookup provinsi
  const provinsiData = wilayahIndonesia[kodeProvinsi];
  if (provinsiData) {
    result.provinsi = provinsiData.nama;

    // Lookup kabupaten
    const kabupatenNama = provinsiData.kabupaten[kodeKabupaten];
    if (kabupatenNama) {
      result.kabupaten = kabupatenNama;
    } else {
      errors.push(`Kode kabupaten/kota "${kodeKabupaten}" tidak ditemukan`);
    }
  } else {
    errors.push(`Kode provinsi "${kodeProvinsi}" tidak ditemukan`);
  }

  // Lookup kecamatan
  const kecamatanNama = kecamatanDatabase[kodeKecamatan];
  if (kecamatanNama) {
    result.kecamatan = kecamatanNama;
  }

  // Parse tanggal lahir & jenis kelamin
  if (cleanNIK.length >= 12) {
    let tanggal = parseInt(cleanNIK.substring(6, 8));
    const bulan = parseInt(cleanNIK.substring(8, 10));
    const tahun = parseInt(cleanNIK.substring(10, 12));

    // Jika tanggal > 40, berarti perempuan (tanggal + 40)
    if (tanggal > 40) {
      result.jenisKelamin = "Perempuan";
      tanggal -= 40;
    } else {
      result.jenisKelamin = "Laki-laki";
    }

    // Determine full year
    const currentYear = new Date().getFullYear() % 100;
    let fullYear: number;
    if (tahun <= currentYear) {
      fullYear = 2000 + tahun;
    } else {
      fullYear = 1900 + tahun;
    }

    // Validate date
    if (bulan < 1 || bulan > 12) {
      errors.push(`Bulan "${bulan}" tidak valid`);
    } else if (tanggal < 1 || tanggal > 31) {
      errors.push(`Tanggal "${tanggal}" tidak valid`);
    } else {
      const dateStr = `${tanggal.toString().padStart(2, "0")}-${bulan.toString().padStart(2, "0")}-${fullYear}`;
      result.tanggalLahir = dateStr;

      // Calculate age
      const today = new Date();
      const birthDate = new Date(fullYear, bulan - 1, tanggal);
      let age = today.getFullYear() - birthDate.getFullYear();
      const monthDiff = today.getMonth() - birthDate.getMonth();
      if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
        age--;
      }
      result.usia = age;
    }
  }

  // Parse urutan registrasi
  if (cleanNIK.length >= 16) {
    result.urutanRegistrasi = cleanNIK.substring(12, 16);
  }

  // Check database penduduk
  const penduduk = database.find((p) => p.nik === cleanNIK);
  if (penduduk) {
    result.penduduk = penduduk;
  }

  result.errors = errors;
  result.valid = errors.length === 0 && cleanNIK.length === 16;

  return result;
}

export function searchByName(nama: string, database: PendudukData[]): PendudukData[] {
  const lowerNama = nama.toLowerCase();
  return database.filter((p) => p.nama.toLowerCase().includes(lowerNama));
}

export function getStatistics(database: PendudukData[]) {
  const total = database.length;

  // Count by Jenis Perolehan
  const jenisPerolehanCount: Record<string, number> = {};
  database.forEach((p) => {
    if (p.jenisPerolehan) {
      jenisPerolehanCount[p.jenisPerolehan] = (jenisPerolehanCount[p.jenisPerolehan] || 0) + 1;
    }
  });

  // Count by Tahun
  const tahunCount: Record<string, number> = {};
  database.forEach((p) => {
    if (p.tahun) {
      tahunCount[p.tahun] = (tahunCount[p.tahun] || 0) + 1;
    }
  });

  // Count by Kecamatan
  const kecamatanCount: Record<string, number> = {};
  database.forEach((p) => {
    if (p.namaKecamatan) {
      kecamatanCount[p.namaKecamatan] = (kecamatanCount[p.namaKecamatan] || 0) + 1;
    }
  });

  // Count by Bulan
  const bulanCount: Record<string, number> = {};
  database.forEach((p) => {
    if (p.bulan) {
      const bulanName = getBulanName(p.bulan);
      bulanCount[bulanName] = (bulanCount[bulanName] || 0) + 1;
    }
  });

  return {
    total,
    jenisPerolehanCount,
    tahunCount,
    kecamatanCount,
    bulanCount,
  };
}

export function getBulanName(bulan: string): string {
  const bulanNames: Record<string, string> = {
    "01": "Januari",
    "1": "Januari",
    "02": "Februari",
    "2": "Februari",
    "03": "Maret",
    "3": "Maret",
    "04": "April",
    "4": "April",
    "05": "Mei",
    "5": "Mei",
    "06": "Juni",
    "6": "Juni",
    "07": "Juli",
    "7": "Juli",
    "08": "Agustus",
    "8": "Agustus",
    "09": "September",
    "9": "September",
    "10": "Oktober",
    "11": "November",
    "12": "Desember",
  };
  return bulanNames[bulan] || bulan;
}

// ============================================
// EXCEL IMPORT / EXPORT
// ============================================

// Normalize header name for flexible column matching
function normalizeHeader(header: string): string {
  return header
    .toString()
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]/g, "");
}

// Map of possible header names to our field names
const headerMap: Record<string, string> = {
  // NIK
  nik: "nik",
  nomorinduk: "nik",
  nomorindukkependudukan: "nik",
  nonik: "nik",
  // Nama
  nama: "nama",
  namalengkap: "nama",
  name: "nama",
  // Jenis Perolehan
  jenisperolehan: "jenisPerolehan",
  jnsperolehan: "jenisPerolehan",
  perolehan: "jenisPerolehan",
  // NPOTKP
  npotkp: "npotkp",
  nonpotkp: "npotkp",
  nomornpotkp: "npotkp",
  // Tahun
  tahun: "tahun",
  thn: "tahun",
  year: "tahun",
  // Bulan
  bulan: "bulan",
  bln: "bulan",
  month: "bulan",
  // Tanggal
  tanggal: "tanggal",
  tgl: "tanggal",
  date: "tanggal",
  // Nama Kecamatan
  namakecamatan: "namaKecamatan",
  kecamatan: "namaKecamatan",
  kec: "namaKecamatan",
  // ID / Nomor
  id: "id",
  no: "id",
  nomor: "id",
  // Keterangan
  keterangan: "keterangan",
  ket: "keterangan",
  catatan: "keterangan",
  notes: "keterangan",
  status: "keterangan",
};

function resolveHeader(raw: string): string | null {
  const normalized = normalizeHeader(raw);
  return headerMap[normalized] || null;
}

/**
 * Parse an Excel file (ArrayBuffer) into PendudukData[]
 */
export function parseExcelBuffer(buffer: ArrayBuffer): { data: PendudukData[]; sheetName: string; totalRows: number; errors: string[] } {
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert to JSON rows – raw values, header row as keys
  const rawRows: Record<string, unknown>[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
  const totalRows = rawRows.length;
  const errors: string[] = [];

  if (totalRows === 0) {
    errors.push("File Excel kosong atau tidak memiliki data.");
    return { data: [], sheetName, totalRows: 0, errors };
  }

  // Detect column mapping from the first row's keys
  const sampleKeys = Object.keys(rawRows[0]);
  const colMap: Record<string, string> = {}; // excelColName -> ourFieldName
  const mappedFields = new Set<string>();

  for (const key of sampleKeys) {
    const field = resolveHeader(key);
    if (field && !mappedFields.has(field)) {
      colMap[key] = field;
      mappedFields.add(field);
    }
  }

  if (!mappedFields.has("nik")) {
    errors.push('Kolom "NIK" tidak ditemukan di file Excel. Pastikan ada kolom bernama NIK.');
    return { data: [], sheetName, totalRows, errors };
  }

  const result: PendudukData[] = [];

  for (let i = 0; i < rawRows.length; i++) {
    const raw = rawRows[i];
    const mapped: Record<string, string> = {};

    for (const [excelCol, fieldName] of Object.entries(colMap)) {
      const rawVal = raw[excelCol];
      let val: string;
      // Convert numbers to string, preserving leading zeros for NIK
      if (typeof rawVal === "number") {
        if (fieldName === "nik") {
          val = String(rawVal).padStart(16, "0");
        } else {
          val = String(rawVal);
        }
      } else {
        val = String(rawVal ?? "").trim();
      }
      mapped[fieldName] = val;
    }

    const nikValue = mapped.nik || "";

    if (!nikValue || nikValue === "0".repeat(16)) {
      continue; // skip empty rows
    }

    const penduduk: PendudukData = {
      id: mapped.id || String(Date.now() + i),
      nik: nikValue,
      keterangan: mapped.keterangan || "",
      nama: mapped.nama || "",
      jenisPerolehan: mapped.jenisPerolehan || "",
      npotkp: mapped.npotkp || "",
      tahun: mapped.tahun || "",
      bulan: mapped.bulan ? String(mapped.bulan).padStart(2, "0") : "",
      tanggal: mapped.tanggal ? String(mapped.tanggal).padStart(2, "0") : "",
      namaKecamatan: mapped.namaKecamatan || "",
    };

    result.push(penduduk);
  }

  if (result.length === 0 && totalRows > 0) {
    errors.push("Tidak ada data NIK valid yang ditemukan. Pastikan kolom NIK terisi.");
  }

  return { data: result, sheetName, totalRows, errors };
}

/**
 * Generate and download an Excel template file
 */
export function downloadExcelTemplate(): void {
  const templateData = [
    {
      No: 1,
      NIK: "3171021505900001",
      Keterangan: "Data Valid",
      Nama: "Contoh Nama Lengkap",
      "Jenis Perolehan": "Pendaftaran Baru",
      NPOTKP: "NPO/3171/2024/001234",
      Tahun: 2024,
      Bulan: 1,
      Tanggal: 15,
      "Nama Kecamatan": "Menteng",
    },
    {
      No: 2,
      NIK: "3273015512850002",
      Keterangan: "Perlu Verifikasi",
      Nama: "Nama Lainnya",
      "Jenis Perolehan": "Pemutakhiran Data",
      NPOTKP: "NPO/3273/2024/002345",
      Tahun: 2024,
      Bulan: 2,
      Tanggal: 20,
      "Nama Kecamatan": "Coblong",
    },
    {
      No: 3,
      NIK: "3578011203880003",
      Keterangan: "KTP-el Diterbitkan",
      Nama: "Contoh Ketiga",
      "Jenis Perolehan": "Penerbitan KTP-el",
      NPOTKP: "NPO/3578/2024/003456",
      Tahun: 2024,
      Bulan: 3,
      Tanggal: 10,
      "Nama Kecamatan": "Wonokromo",
    },
  ];

  const worksheet = XLSX.utils.json_to_sheet(templateData);

  // Set column widths
  worksheet["!cols"] = [
    { wch: 5 },  // No
    { wch: 20 }, // NIK
    { wch: 20 }, // Keterangan
    { wch: 25 }, // Nama
    { wch: 22 }, // Jenis Perolehan
    { wch: 24 }, // NPOTKP
    { wch: 8 },  // Tahun
    { wch: 8 },  // Bulan
    { wch: 8 },  // Tanggal
    { wch: 20 }, // Nama Kecamatan
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Data NIK");

  // Write and download
  XLSX.writeFile(workbook, "template_import_nik.xlsx");
}

/**
 * Export current database to Excel file
 */
export function exportDatabaseToExcel(database: PendudukData[]): void {
  const exportData = database.map((p, i) => ({
    No: i + 1,
    NIK: p.nik,
    Keterangan: p.keterangan,
    Nama: p.nama,
    "Jenis Perolehan": p.jenisPerolehan,
    NPOTKP: p.npotkp,
    Tahun: p.tahun,
    Bulan: p.bulan,
    Tanggal: p.tanggal,
    "Nama Kecamatan": p.namaKecamatan,
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);

  // Set column widths
  worksheet["!cols"] = [
    { wch: 5 },  // No
    { wch: 20 }, // NIK
    { wch: 20 }, // Keterangan
    { wch: 25 }, // Nama
    { wch: 22 }, // Jenis Perolehan
    { wch: 24 }, // NPOTKP
    { wch: 8 },  // Tahun
    { wch: 8 },  // Bulan
    { wch: 8 },  // Tanggal
    { wch: 20 }, // Nama Kecamatan
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Data NIK");

  XLSX.writeFile(workbook, `database_nik_${new Date().toISOString().slice(0, 10)}.xlsx`);
}
