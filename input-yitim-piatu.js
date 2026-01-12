/**
 * input-yatim-piatu.js
 *
 * Otomasi input warga terdata (Suplemen Yatim Piatu):
 * URL: http://opensid.local/suplemen/form_terdata/2/1
 *
 * - Pilih Warga Terdata via select2 #terdata (search by NAMA)
 * - Validasi pilihan harus mengandung NIK dari Excel
 * - Fallback: kalau search nama gagal -> search pakai NIK
 * - Submit form (Simpan)
 * - Buat report Excel: report-yatim-piatu.xlsx (baris gagal kuning + kolom PESAN_GAGAL)
 * - Reprocess: baris gagal diulang max 3x di akhir
 */

const { chromium } = require("playwright");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const BASE = "http://opensid.local";
const LOGIN_URL = `${BASE}/siteman`;
const FORM_URL = `${BASE}/suplemen/form_terdata/2/1`;

const EXCEL_PATH = path.resolve(__dirname, "DATA YATIM TAHUN 2023.xlsx");
const STORAGE_STATE_PATH = path.resolve(__dirname, "storageState.json");
const REPORT_PATH = path.resolve(__dirname, "report-yatim-piatu.xlsx");

const USERNAME = process.env.OPENSID_USER || "admin";
const PASSWORD = process.env.OPENSID_PASS || "Admin123.";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function cleanText(s) {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();
}
function upper(s) {
  return cleanText(s).toUpperCase();
}
function onlyDigits(s) {
  return String(s ?? "").replace(/\D/g, "");
}
function escapeRegExp(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Minimal normalisasi nama (biar tahan prefix)
function coreName(nama) {
  let s = cleanText(nama);
  s = s.replace(/^\s*R\s*\.\s*/i, "");
  s = s.replace(/^\s*RADEN\s+/i, "");
  return cleanText(s);
}
function buildNameQueries(namaRaw) {
  const base = cleanText(namaRaw);
  const core = coreName(base);
  const q = [];
  if (base) q.push(base);
  if (core && core !== base) {
    q.push(`Raden ${core}`);
    q.push(core);
  }
  return [...new Set(q)];
}

async function ensureLoggedIn(page) {
  await page.goto(LOGIN_URL, { waitUntil: "domcontentloaded" });

  const usernameInput = page.locator('input[name="username"]');
  const needsLogin = await usernameInput.isVisible().catch(() => false);
  if (!needsLogin) return;

  await usernameInput.fill(USERNAME);
  await page.locator("#password").fill(PASSWORD);

  await Promise.all([
    page.waitForLoadState("networkidle").catch(() => {}),
    page.locator('button[type="submit"]', { hasText: "Masuk" }).click(),
  ]);

  const stillOnLogin = await usernameInput.isVisible().catch(() => false);
  if (stillOnLogin) {
    const errText = await page
      .locator(".alert, .error, .text-danger")
      .first()
      .textContent()
      .catch(() => "");
    throw new Error(
      `Login gagal / masih di halaman login. ${
        errText ? "Pesan: " + cleanText(errText) : ""
      }`
    );
  }
}

/**
 * Baca Excel: cari header row yang mengandung NIK dan NAMA.
 * Output object mempertahankan kolom asli, tapi minimal ada field: NIK, NAMA (atau yang terdeteksi).
 */
function readYatimData(filePath) {
  const wb = xlsx.readFile(filePath);
  const sheetName = wb.SheetNames[0];
  const sh = wb.Sheets[sheetName];

  const aoa = xlsx.utils.sheet_to_json(sh, { header: 1, defval: "" });

  const norm = (v) => upper(v).replace(/\./g, "").replace(/\s+/g, " ").trim();

  // cari header row (maks 50 baris awal biar cepat)
  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(50, aoa.length); i++) {
    const row = aoa[i] || [];
    const cells = row.map(norm);

    const hasNik = cells.some((c) => c === "NIK" || c.includes("NIK"));
    const hasNama = cells.some(
      (c) => c === "NAMA" || c.includes("NAMA") || c.includes("NAMA LENGKAP")
    );

    if (hasNik && hasNama) {
      headerRowIndex = i;
      break;
    }
  }

  // fallback kalau tidak ketemu: pakai sheet_to_json biasa
  if (headerRowIndex < 0) {
    const rows = xlsx.utils.sheet_to_json(sh, { defval: "" });
    return rows;
  }

  const header = (aoa[headerRowIndex] || []).map((h) => cleanText(h));
  const dataRows = aoa.slice(headerRowIndex + 1);

  // cari index kolom NIK & NAMA
  const headerNorm = header.map(norm);
  const idxNik = headerNorm.findIndex((h) => h === "NIK" || h.includes("NIK"));
  const idxNama = headerNorm.findIndex(
    (h) => h === "NAMA" || h.includes("NAMA")
  );

  const out = [];
  for (const r of dataRows) {
    const row = r || [];

    // buat object semua kolom biar "data awal" tetap ada di report
    const obj = {};
    for (let c = 0; c < header.length; c++) {
      const key = header[c] || `COL_${c + 1}`;
      obj[key] = row[c] ?? "";
    }

    // tambahkan akses field standar
    obj.__NIK = idxNik >= 0 ? onlyDigits(row[idxNik]) : onlyDigits(obj["NIK"]);
    obj.__NAMA =
      idxNama >= 0 ? cleanText(row[idxNama]) : cleanText(obj["NAMA"]);

    // skip benar-benar kosong
    if (!cleanText(obj.__NAMA) && !cleanText(obj.__NIK)) continue;

    out.push(obj);
  }

  return out;
}

/**
 * Select2: pilih warga terdata (id terdata)
 * - cari pakai nama (variasi) -> validasi harus mengandung NIK
 * - fallback cari pakai NIK
 */
async function pickTerdataByNameValidateNik(page, namaRaw, nikRaw) {
  const nama = cleanText(namaRaw);
  const nik = onlyDigits(nikRaw);

  if (!nik) throw new Error("NIK kosong / tidak valid");
  if (!nama) throw new Error("Nama kosong");

  const nameQueries = buildNameQueries(nama);
  const queries = [...nameQueries, nik]; // fallback by nik

  const nikRegex = new RegExp(escapeRegExp(nik), "i");
  const core = coreName(nama);
  const coreRegex = core ? new RegExp(escapeRegExp(core), "i") : null;

  const select2Clickable = page
    .locator(
      [
        "#select2-terdata-container",
        'span.select2-selection--single[aria-labelledby="select2-terdata-container"]',
        "span.select2-container:has(#select2-terdata-container)",
      ].join(",")
    )
    .first();

  await select2Clickable.waitFor({ state: "visible", timeout: 15000 });

  const MAX_ROUNDS = 3;
  const TYPE_DELAY = 160;
  const STABILIZE_WAIT = 1100;

  for (let round = 1; round <= MAX_ROUNDS; round++) {
    for (const q of queries) {
      // open dropdown
      const openCount = await page.locator(".select2-container--open").count();
      if (!openCount) await select2Clickable.click();

      const searchField = page.locator("input.select2-search__field").first();
      await searchField.waitFor({ state: "visible", timeout: 10000 });

      await searchField.fill("");
      await sleep(150);
      await searchField.type(String(q), { delay: TYPE_DELAY });

      // tunggu ajax/stabilisasi
      await Promise.race([
        page
          .waitForResponse(
            (res) =>
              res.ok() &&
              /select2|penduduk|autocomplete|ajax|cari|suplemen/i.test(
                res.url()
              ) &&
              (res.request().method() === "GET" ||
                res.request().method() === "POST"),
            { timeout: 6500 }
          )
          .catch(() => null),
        sleep(STABILIZE_WAIT),
      ]);

      await page.waitForSelector(".select2-results__option", {
        timeout: 12000,
      });

      // no results -> lanjut query berikutnya
      const noResult = page.locator(".select2-results__message").first();
      if (await noResult.isVisible().catch(() => false)) continue;

      // pilih opsi yang mengandung NIK (paling valid)
      const optionByNik = page
        .locator(".select2-results__option", { hasText: nikRegex })
        .filter({ hasNot: page.locator(".select2-results__message") })
        .first();

      let picked = false;

      if ((await optionByNik.count()) > 0) {
        await optionByNik.click();
        picked = true;
      } else if (coreRegex) {
        // fallback: kalau dropdown tidak menampilkan nik, cari yg mengandung nama inti
        const optionByName = page
          .locator(".select2-results__option", { hasText: coreRegex })
          .filter({ hasNot: page.locator(".select2-results__message") })
          .first();

        if ((await optionByName.count()) > 0) {
          await optionByName.click();
          picked = true;
        }
      }

      if (!picked) continue;

      // tunggu dropdown menutup
      await page
        .locator(".select2-container--open")
        .waitFor({ state: "detached", timeout: 6000 })
        .catch(() => {});

      // validasi selection: harus mengandung NIK
      const selectedText = cleanText(
        await page.locator("#select2-terdata-container").textContent()
      );
      const selectedNik = onlyDigits(selectedText);

      if (!selectedNik.includes(nik)) {
        // coba round berikutnya / query berikutnya
        continue;
      }

      return; // sukses
    }

    await sleep(400);
  }

  throw new Error(
    `Tidak menemukan warga terdata yang cocok untuk NIK=${nik} | Nama=${nama}`
  );
}

async function writeReportExcel(originalRows, errorMessages, outPath) {
  // report = semua kolom asli (kecuali field internal __NIK/__NAMA) + PESAN_GAGAL
  const cleaned = originalRows.map((r) => {
    const copy = { ...r };
    delete copy.__NIK;
    delete copy.__NAMA;
    return copy;
  });

  const originalKeys = Array.from(
    new Set(cleaned.flatMap((r) => Object.keys(r)))
  );
  const headers = [...originalKeys, "PESAN_GAGAL"];

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("REPORT");

  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };

  const yellowFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF00" },
  };

  for (let i = 0; i < cleaned.length; i++) {
    const base = cleaned[i];
    const pesan = cleanText(errorMessages[i] || "");

    const rowValues = headers.map((h) =>
      h === "PESAN_GAGAL" ? pesan : base[h] ?? ""
    );
    const row = ws.addRow(rowValues);

    if (pesan) {
      row.eachCell((cell) => {
        cell.fill = yellowFill;
      });
    }
  }

  ws.columns.forEach((col) => {
    const headerLen = String(col.header || "").length;
    col.width = Math.min(60, Math.max(12, headerLen + 2));
  });

  await wb.xlsx.writeFile(outPath);
}

// ====== proses 1 baris (dipakai pass 1 & reprocess) ======
async function processOneRow(page, context, row) {
  const nama = cleanText(row.__NAMA ?? row.NAMA ?? row.nama ?? "");
  const nik = onlyDigits(row.__NIK ?? row.NIK ?? row.nik ?? "");

  if (!nama) return "Nama kosong";
  if (!nik) return "NIK kosong";

  await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });

  // kalau kelempar login lagi
  if (page.url().includes("/siteman")) {
    await ensureLoggedIn(page);
    await context.storageState({ path: STORAGE_STATE_PATH });
    await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });
  }

  await pickTerdataByNameValidateNik(page, nama, nik);

  // keterangan: biarkan kosong (kalau mau isi default, uncomment)
  // await page.locator('textarea[name="keterangan"]').fill("-");

  const simpanBtn = page.locator('button[type="submit"]', {
    hasText: /Simpan/i,
  });
  await simpanBtn.waitFor({ state: "visible", timeout: 15000 });

  await Promise.allSettled([
    page.waitForLoadState("networkidle", { timeout: 15000 }),
    simpanBtn.click(),
  ]);

  return ""; // sukses
}

// ====== reprocess baris gagal max 3x ======
async function reprocessFailed(
  page,
  context,
  rows,
  errorMessages,
  maxRetry = 3
) {
  for (let attempt = 1; attempt <= maxRetry; attempt++) {
    const failedIdx = [];
    for (let i = 0; i < errorMessages.length; i++) {
      if (cleanText(errorMessages[i])) failedIdx.push(i);
    }
    if (failedIdx.length === 0) return;

    console.log(
      `\n=== REPROCESS attempt ${attempt}/${maxRetry} | gagal: ${failedIdx.length} baris ===`
    );
    await sleep(800);

    for (const idx of failedIdx) {
      const r = rows[idx];
      const nama = cleanText(r.__NAMA ?? r.NAMA ?? r.nama ?? "");
      const nik = onlyDigits(r.__NIK ?? r.NIK ?? r.nik ?? "");

      console.log(
        `  [RETRY ${attempt}] idx=${idx + 1} | NIK=${nik} | Nama=${nama}`
      );

      try {
        const err = await processOneRow(page, context, r);
        if (!err) {
          errorMessages[idx] = "";
          console.log("    - Sukses (reprocess)");
        } else {
          errorMessages[idx] = cleanText(err);
          console.log("    - Masih gagal:", errorMessages[idx]);
        }
      } catch (e) {
        errorMessages[idx] = cleanText(e.message);
        console.log("    - Exception:", errorMessages[idx]);
      }

      await sleep(300);
    }
  }
}

async function main() {
  const rows = readYatimData(EXCEL_PATH);
  const errorMessages = rows.map(() => "");

  const browser = await chromium.launch({ headless: false, slowMo: 50 });
  const context = await browser.newContext(
    fs.existsSync(STORAGE_STATE_PATH)
      ? { storageState: STORAGE_STATE_PATH }
      : {}
  );
  const page = await context.newPage();

  try {
    await ensureLoggedIn(page);
    await context.storageState({ path: STORAGE_STATE_PATH });
    console.log("Login OK (session tersimpan).");
  } catch (e) {
    await browser.close();
    throw e;
  }

  // PASS 1
  for (let idx = 0; idx < rows.length; idx++) {
    const r = rows[idx];
    const nama = cleanText(r.__NAMA ?? r.NAMA ?? r.nama ?? "");
    const nik = onlyDigits(r.__NIK ?? r.NIK ?? r.nik ?? "");

    console.log(`\n[${idx + 1}/${rows.length}] NIK=${nik} | Nama=${nama}`);

    try {
      const err = await processOneRow(page, context, r);
      errorMessages[idx] = cleanText(err);
      console.log(err ? `  - Gagal: ${errorMessages[idx]}` : "  - Sukses");
    } catch (e) {
      errorMessages[idx] = cleanText(e.message);
      console.log("  - Gagal:", errorMessages[idx]);
    }

    await sleep(250);
  }

  // REPROCESS gagal max 3x
  await reprocessFailed(page, context, rows, errorMessages, 3);

  await browser.close();

  await writeReportExcel(rows, errorMessages, REPORT_PATH);
  console.log(`\nSelesai. Report dibuat: ${REPORT_PATH}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
