function processExcel() {
  const file = document.getElementById("fileInput").files[0];
  if (!file) return alert("Pilih file Excel dulu");

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    // 1️⃣ URUTKAN DATA
    rows.sort((a, b) => {
      const ta = new Date(`${a["Tanggal Absensi"]}T${a["Jam Absensi"]}`);
      const tb = new Date(`${b["Tanggal Absensi"]}T${b["Jam Absensi"]}`);
      return ta - tb;
    });

    const sesiAktif = {};
    const sheets = {};
    const summary = {};

    rows.forEach(r => {
      const ID = r["ID"];
      const nama = r["Nama"];
      const jabatan = r["Jabatan"];
      const tanggal = r["Tanggal Absensi"];
      const jam = r["Jam Absensi"];
      const tipe = (r["Tipe Absensi"] || "").toLowerCase();

      if (!nama || !tanggal || !jam) return;

      if (!sesiAktif[nama]) sesiAktif[nama] = null;

      // ===== MASUK =====
      if (tipe.includes("masuk")) {
        sesiAktif[nama] = {
          ID,
          nama,
          jabatan,
          tglMasuk: tanggal,
          jamMasuk: jam
        };
      }

      // ===== PULANG =====
      if (tipe.includes("pulang")) {
        const sheetName = tentukanSheet(jabatan);
        if (!sheets[sheetName]) sheets[sheetName] = [];

        let hasil;
        let tglMasuk = "";
        let jamMasuk = "";

        if (sesiAktif[nama]) {
          // MASUK + PULANG LENGKAP
          tglMasuk = sesiAktif[nama].tglMasuk;
          jamMasuk = sesiAktif[nama].jamMasuk;
          hasil = hitungJamKerja(jamMasuk, jam, tglMasuk);
        } else {
          // PULANG TANPA MASUK → 0,5
          hasil = { hariKerja: 0.5, lembur: 0 };
        }

        sheets[sheetName].push({
          ID: ID,
          Nama: nama,
          Jabatan: jabatan,
          Tanggal_Masuk: tglMasuk,
          Jam_Masuk: jamMasuk,
          Tanggal_Pulang: tanggal,
          Jam_Pulang: jam,
          Hari_Kerja: hasil.hariKerja,
          Jam_Lembur: hasil.lembur.toFixed(2)
        });

        if (!summary[nama]) {
          summary[nama] = {
            ID: ID,
            Nama: nama,
            Jabatan: jabatan,
            Total_Hari_Kerja: 0,
            Total_Jam_Lembur: 0
          };
        }

        summary[nama].Total_Hari_Kerja += hasil.hariKerja;
        summary[nama].Total_Jam_Lembur += hasil.lembur;

        sesiAktif[nama] = null;
      }
    });

    // 2️⃣ HANDLE MASUK TANPA PULANG (SESI GANTUNG)
    Object.values(sesiAktif).forEach(sesi => {
      if (!sesi) return;

      const sheetName = tentukanSheet(sesi.jabatan);
      if (!sheets[sheetName]) sheets[sheetName] = [];

      sheets[sheetName].push({
        ID: sesi.ID,
        Nama: sesi.nama,
        Jabatan: sesi.jabatan,
        Tanggal_Masuk: sesi.tglMasuk,
        Jam_Masuk: sesi.jamMasuk,
        Tanggal_Pulang: "",
        Jam_Pulang: "",
        Hari_Kerja: 0.5,
        Jam_Lembur: "0.00"
      });

      if (!summary[sesi.nama]) {
        summary[sesi.nama] = {
          ID: sesi.ID,
          Nama: sesi.nama,
          Jabatan: sesi.jabatan,
          Total_Hari_Kerja: 0,
          Total_Jam_Lembur: 0
        };
      }

      summary[sesi.nama].Total_Hari_Kerja += 0.5;
    });

    // 3️⃣ EXPORT
    const outWB = XLSX.utils.book_new();

    Object.keys(sheets).forEach(name => {
      const wsOut = XLSX.utils.json_to_sheet(sheets[name]);
      XLSX.utils.book_append_sheet(outWB, wsOut, name.substring(0, 31));
    });

    const wsSummary = XLSX.utils.json_to_sheet(
      Object.values(summary).map(s => ({
        ID: s.ID,
        Nama: s.Nama,
        Jabatan: s.Jabatan,
        Total_Hari_Kerja: s.Total_Hari_Kerja,
        Total_Jam_Lembur: s.Total_Jam_Lembur.toFixed(2)
      }))
    );

    XLSX.utils.book_append_sheet(outWB, wsSummary, "SUMMARY");
    XLSX.writeFile(outWB, "HASIL_PAYROLL_FINAL.xlsx");
  };

  reader.readAsArrayBuffer(file);
}

/* =====================================================
   HITUNG JAM KERJA FINAL (SESUAI ATURAN)
===================================================== */
function hitungJamKerja(jamMasuk, jamPulang, tanggal) {

  function normalisasiJamMasuk(jam) {
    if (!jam) return null;
    let [h, m] = jam.split(":").map(Number);
    const shiftJam = [7, 15, 23];
    for (let sj of shiftJam) {
      if (h === sj && m <= 29) return String(sj).padStart(2, "0") + ":00";
      if (h === sj - 1 && m > 29) return String(sj).padStart(2, "0") + ":00";
    }
    return String(h).padStart(2, "0") + ":" + String(m).padStart(2, "0");
  }

  function normalisasiJamPulang(jam) {
    if (!jam) return null;
    let [h, m] = jam.split(":").map(Number);
    m = m < 30 ? 0 : 30;
    return String(h).padStart(2, "0") + ":" + String(m).padStart(2, "0");
  }

  jamMasuk = normalisasiJamMasuk(jamMasuk);
  jamPulang = normalisasiJamPulang(jamPulang);

  // Data parsial → hanya masuk/pulang
  if (!jamMasuk || !jamPulang) return { hariKerja: 0.5, lembur: 0 };

  let masuk = new Date(`${tanggal}T${jamMasuk}`);
  let pulang = new Date(`${tanggal}T${jamPulang}`);

  // SHIFT MALAM / LINTAS HARI
  if (jamMasuk >= "23:00" && jamPulang < jamMasuk) {
    pulang.setDate(pulang.getDate() + 1);
  } else if (pulang <= masuk) {
    pulang.setDate(pulang.getDate() + 1);
  }

  const hari = masuk.getDay();
  let hariKerja = 0;
  let lembur = 0;

  // MINGGU
  if (hari === 0) {
    let durasi = (pulang - masuk) / 3600000 - 1;
    return { hariKerja: 0, lembur: Math.max(0, durasi) };
  }

  // SABTU
  if (hari === 6) {
    hariKerja = 1;
    let mulaiLembur = new Date(masuk);
    mulaiLembur.setMinutes(mulaiLembur.getMinutes() + (5.5 * 60) + 60);
    lembur = pulang > mulaiLembur ? (pulang - mulaiLembur) / 3600000 : 0;
    return { hariKerja, lembur };
  }

  // SENIN-JUMAT
  hariKerja = 1;
  let durasi = (pulang - masuk) / 3600000 - 1;
  lembur = Math.max(0, durasi - 7);
  return { hariKerja, lembur };
}


function tentukanSheet(jabatan) {
  if (!jabatan) return "TANPA_JABATAN";
  let j = jabatan.toUpperCase();

  if (j.includes("SATPAM") || j.includes("SECURITY")) return "SATPAM";
  if (j.includes("QC") || j.includes("KOORDINATOR QC")) return "QC";
  if (j.includes("A ASSEMBLY") ||
      j.includes("KARU 1 ASSEMBLY") ||
      j.includes("KARU 2 ASSEMBLY") ||
      j.includes("KARU 3 ASSEMBLY")) return "ASSEMBLY";
  if (j.includes("INJECT") && j.includes("1")) return "INJECT 1";
  if (j.includes("INJECT") && j.includes("2")) return "INJECT 2";
  if (j.includes("INJECT") && j.includes("3")) return "INJECT 3";
  if (j.includes("SPRAY") || j.includes("KARU SPRAY") ||
      j.includes("PRINTING") || j.includes("TU") ||
      j.includes("ADMIN SPRAY PRINTING") || j.includes("KARU TU") ||
      j.includes("KARU PRINTING") || j.includes("SPV SPRAY") ||
      j.includes("SETTING CAT")) return "SPRAY PRINTING";
  if (["MANAGER","MANAGER HRD","ACCOUNTING","TRANSLATOR","PPIC",
       "ADMIN WAREHOUSE","EXIM","HSE","HRD"].includes(j)) return "OFFICE";

  return j.replace(/[:\\\/\?\*\[\]]/g, "").replace(/\s+/g, "_").substring(0,31);
}
