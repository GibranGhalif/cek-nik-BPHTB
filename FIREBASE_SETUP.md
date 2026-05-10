# 🔥 Setup Firebase untuk Database Online Real-time

Aplikasi ini mendukung penyimpanan data secara **online dan real-time** menggunakan Firebase Firestore. Ikuti langkah-langkah berikut untuk mengaktifkan fitur ini:

## 📋 Langkah-langkah Setup

### 1. Buat Project Firebase

1. Buka [Firebase Console](https://console.firebase.google.com/)
2. Klik **"Add project"** atau **"Tambah project"**
3. Masukkan nama project (contoh: `cek-nik-app`)
4. (Opsional) Nonaktifkan Google Analytics jika tidak diperlukan
5. Klik **"Create project"**

### 2. Tambahkan Web App

1. Di dashboard project, klik ikon **Web** (</>) untuk menambahkan aplikasi web
2. Masukkan nickname aplikasi (contoh: `cek-nik-web`)
3. **Jangan** centang "Firebase Hosting" (tidak diperlukan)
4. Klik **"Register app"**
5. Anda akan melihat konfigurasi Firebase seperti ini:

```javascript
const firebaseConfig = {
  apiKey: "AIzaSy...",
  authDomain: "your-project-id.firebaseapp.com",
  projectId: "your-project-id",
  storageBucket: "your-project-id.appspot.com",
  messagingSenderId: "123456789",
  appId: "1:123456789:web:abcdef123456"
};
```

### 3. Aktifkan Firestore Database

1. Di sidebar Firebase Console, klik **"Build"** → **"Firestore Database"**
2. Klik **"Create database"**
3. Pilih **"Start in test mode"** (untuk development)
   - ⚠️ Untuk production, ubah rules ke mode secured
4. Pilih lokasi server terdekat (contoh: `asia-southeast2` untuk Indonesia)
5. Klik **"Done"**

### 4. Update Konfigurasi di Aplikasi

Buka file `src/lib/firebase.ts` dan ganti konfigurasi placeholder dengan konfigurasi dari Firebase Console:

```typescript
const firebaseConfig = {
  apiKey: "AIzaSy...",           // Ganti dengan API Key Anda
  authDomain: "your-project-id.firebaseapp.com",
  projectId: "your-project-id",   // Ganti dengan Project ID Anda
  storageBucket: "your-project-id.appspot.com",
  messagingSenderId: "123456789",
  appId: "1:123456789:web:abcdef123456"
};
```

### 5. (Opsional) Setup Security Rules untuk Production

Untuk keamanan di production, ubah Firestore Rules di Firebase Console:

```javascript
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /penduduk/{document=**} {
      // Izinkan read untuk semua
      allow read: if true;
      
      // Izinkan write hanya jika data valid
      allow write: if request.resource.data.nik is string 
                   && request.resource.data.nik.size() == 16;
    }
  }
}
```

---

## 🎮 Cara Penggunaan

### Mode Offline (Default)
- Data disimpan di **localStorage** browser
- Tidak perlu koneksi internet
- Data hanya tersedia di browser/device yang sama

### Mode Online
- Data disimpan di **Firebase Firestore**
- Sinkronisasi **real-time** antar semua device
- Perlu konfigurasi Firebase terlebih dahulu

### Mengganti Mode
1. Klik tombol **"💾 Local"** atau **"☁️ Cloud"** di header aplikasi
2. Jika Firebase belum dikonfigurasi, akan muncul pesan error

---

## ✨ Fitur Real-time

Ketika mode online aktif:
- ✅ **Auto-sync**: Perubahan otomatis tersinkronisasi ke semua device
- ✅ **Real-time updates**: Data berubah langsung tanpa perlu refresh
- ✅ **Indikator status**: Badge hijau "Real-time" muncul di header database
- ✅ **Timestamp**: Waktu terakhir sinkronisasi ditampilkan

---

## 🔧 Troubleshooting

### "Firebase belum dikonfigurasi"
- Pastikan sudah mengisi konfigurasi di `src/lib/firebase.ts`
- Pastikan `apiKey` dan `projectId` sudah benar

### "Gagal terhubung ke database online"
- Periksa koneksi internet
- Pastikan Firestore Database sudah diaktifkan di Firebase Console
- Cek Security Rules tidak memblokir akses

### Data tidak sinkron
- Periksa console browser untuk error
- Pastikan semua device menggunakan project Firebase yang sama

---

## 📱 Kompatibilitas

- ✅ Chrome, Firefox, Safari, Edge
- ✅ Desktop & Mobile
- ✅ Multiple tabs/windows
- ✅ Multiple devices

---

## 💡 Tips

1. **Development**: Gunakan test mode rules untuk kemudahan
2. **Production**: Selalu gunakan security rules yang ketat
3. **Backup**: Export data ke Excel secara berkala
4. **Monitoring**: Gunakan Firebase Console untuk monitoring usage
