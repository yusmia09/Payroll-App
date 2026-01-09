function processGantiFormat() {
  const file = document.getElementById("fileInput").files[0];
  if (!file) return alert("Pilih file HASIL_PAYROLL_FINAL.xlsx dulu");

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: "array" });

    const sheetsOut = {};
    const wbOut = XLSX.utils.book_new();

    /* ==========================================
       1️⃣ BACA SEMUA SHEET (KECUALI SUMMARY)
    ========================================== */
    wb.SheetNames.forEach(sheetIn => {
      if (sheetIn === "SUMMARY") return;

      const ws = wb.Sheets[sheetIn];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

      rows.forEach(r => {
        const nama = r["Nama"];
        const jabatan = r["Jabatan"];
        const tglMasuk = r["Tanggal_Masuk"];

        // === FIX KOMA DESIMAL (0,5) ===
        const hariKerjaRaw = r["Hari_Kerja"];
        const hariKerja =
          hariKerjaRaw === "" || hariKerjaRaw === null
            ? ""
            : Number(String(hariKerjaRaw).replace(",", "."));

        const lemburRaw = r["Jam_Lembur"];
        const lembur =
          lemburRaw === "" || lemburRaw === null
            ? ""
            : Number(String(lemburRaw).replace(",", "."));

        if (!nama || !tglMasuk) return;

        const d = new Date(tglMasuk).getDate();
        const sheetName = tentukanSheet(jabatan);

        if (!sheetsOut[sheetName]) sheetsOut[sheetName] = {};
        if (!sheetsOut[sheetName][nama]) {
          sheetsOut[sheetName][nama] = {
            Nama: nama,
            Jabatan: jabatan,
            hari: {},
            lembur: {}
          };
        }

        // === ANTI NINDIH (1 ORANG + 1 TANGGAL) ===
        if (sheetsOut[sheetName][nama].hari[d] === undefined) {
          sheetsOut[sheetName][nama].hari[d] = hariKerja;
        }

        if (sheetsOut[sheetName][nama].lembur[d] === undefined) {
          sheetsOut[sheetName][nama].lembur[d] = lembur;
        }
      });
    });

    /* ==========================================
       2️⃣ BENTUK FORMAT BULANAN
    ========================================== */
    Object.keys(sheetsOut).forEach(sheetName => {
      const dataOut = [];

      Object.values(sheetsOut[sheetName]).forEach(row => {
        const out = {};

        out["NAMA"] = row.Nama;
        out["BAGIAN"] = row.Jabatan;

        // === HARI 1–31 (LANGSUNG DI BAWAH TANGGAL) ===
        for (let d = 1; d <= 31; d++) {
          out[d] = row.hari[d] !== undefined ? row.hari[d] : "";
        }

        out["TOTAL HARI KERJA"] =
          Object.values(row.hari)
            .filter(v => v !== "")
            .reduce((a, b) => a + Number(b), 0);

        // === LEMBUR 1–31 ===
        for (let d = 1; d <= 31; d++) {
          out[`L${d}`] = row.lembur[d] !== undefined ? row.lembur[d] : "";
        }

        out["TOTAL LEMBUR"] =
          Object.values(row.lembur)
            .filter(v => v !== "")
            .reduce((a, b) => a + Number(b), 0);

        dataOut.push(out);
      });

      /* ==========================================
         3️⃣ HEADER (TIDAK BOLEH DOBEL)
      ========================================== */
      const header = ["NAMA", "BAGIAN"];

      for (let d = 1; d <= 31; d++) header.push(String(d));

      header.push("TOTAL HARI KERJA");

      for (let d = 1; d <= 31; d++) header.push(`L${d}`);

      header.push("TOTAL LEMBUR");
      const wsOut = XLSX.utils.json_to_sheet(dataOut, {
        header: header,
        skipHeader: false
      });

      /* ==========================================
         4️⃣ FREEZE PANES (NAMA + BAGIAN & HEADER)
      ========================================== */
      wsOut["!freeze"] = { xSplit: 2, ySplit: 1 };

      /* ==========================================
         5️⃣ LEBAR KOLOM
      ========================================== */
      const colWidths = [];
      colWidths.push({ wch: 30 }); // NAMA
      colWidths.push({ wch: 18 }); // BAGIAN
      for (let i = 1; i <= 31; i++) colWidths.push({ wch: 4 });
      colWidths.push({ wch: 16 }); // TOTAL HARI
      for (let i = 1; i <= 31; i++) colWidths.push({ wch: 4 });
      colWidths.push({ wch: 16 }); // TOTAL LEMBUR

      wsOut["!cols"] = colWidths;

      XLSX.utils.book_append_sheet(
        wbOut,
        wsOut,
        sheetName.substring(0, 31)
      );
    });

    XLSX.writeFile(wbOut, "FORMAT_BULANAN_DESEMBER.xlsx");
  };

  reader.readAsArrayBuffer(file);
}

/* =====================================================
   TENTUKAN SHEET
===================================================== */
function tentukanSheet(jabatan) {
  if (!jabatan) return "TANPA_JABATAN";
  let j = jabatan.toUpperCase();

  if (j.includes("SATPAM") || j.includes("SECURITY")) return "SATPAM";
  if (j.includes("QC")) return "QC";

  if (
    j.includes("A ASSEMBLY") ||
    j.includes("KARU 1 ASSEMBLY") ||
    j.includes("KARU 2 ASSEMBLY") ||
    j.includes("KARU 3 ASSEMBLY")
  ) return "ASSEMBLY";

  if (j.includes("INJECT") && j.includes("1")) return "INJECT 1";
  if (j.includes("INJECT") && j.includes("2")) return "INJECT 2";
  if (j.includes("INJECT") && j.includes("3")) return "INJECT 3";

  if (
    j.includes("SPRAY") ||
    j.includes("PRINTING") ||
    j.includes("TU") ||
    j.includes("SETTING CAT")
  ) return "SPRAY PRINTING";

  if (
    ["MANAGER","MANAGER HRD","ACCOUNTING","TRANSLATOR","PPIC",
     "ADMIN WAREHOUSE","EXIM","HSE","HRD"].includes(j)
  ) return "OFFICE";

  return j.replace(/[:\\\/\?\*\[\]]/g, "")
          .replace(/\s+/g, "_")
          .substring(0, 31);
}
