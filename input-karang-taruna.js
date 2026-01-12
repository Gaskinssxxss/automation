/**
 * input-karang-taruna.js (dengan REPROCESS gagal max 3x)
 */

const { chromium } = require("playwright");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const BASE = "http://opensid.local";
const LOGIN_URL = `${BASE}/siteman`;
const FORM_URL = `${BASE}/lembaga_anggota/form/52`;

const EXCEL_PATH = path.resolve(__dirname, "data_pengurus_karang_taruna.xlsx");
const STORAGE_STATE_PATH = path.resolve(__dirname, "storageState.json");
const REPORT_PATH = path.resolve(__dirname, "report-karang-taruna.xlsx");

const USERNAME = process.env.OPENSID_USER || "admin";
const PASSWORD = process.env.OPENSID_PASS || "Admin123.";

// Konstanta sesuai permintaan
const SK_PENGANGKATAN_NO = "141/14/01/2025";
const SK_PENGANGKATAN_TGL = "17-01-2025";
const SK_PEMBERHENTIAN_NO = "-";
const SK_PEMBERHENTIAN_TGL = "10-01-2026";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function cleanText(s) {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();
}
function upper(s) {
  return cleanText(s).toUpperCase();
}
function escapeRegExp(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ====== Nama inti untuk prefix "R." / "Raden" ======
function coreName(nama) {
  let s = cleanText(nama);
  s = s.replace(/^\s*R\s*\.\s*/i, "");
  s = s.replace(/^\s*RADEN\s+/i, "");
  return cleanText(s);
}
function buildQueriesNama(namaRaw) {
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
 * Baca Excel:
 * Header kemungkinan: no, nama, Jabatan/Bidang (atau jabatan/bidang)
 * Kalau header tidak baris pertama, cari baris yang punya minimal: NO + NAMA
 */
function readKarangTarunaData(filePath) {
  const wb = xlsx.readFile(filePath);
  const sh = wb.Sheets[wb.SheetNames[0]];
  const aoa = xlsx.utils.sheet_to_json(sh, { header: 1, defval: "" });

  const norm = (v) => upper(v).replace(/\./g, "").replace(/\s+/g, " ").trim();

  const findHeaderRowIndex = () => {
    for (let i = 0; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const cells = row.map(norm);
      const hasNo = cells.some((c) => c === "NO" || c === "NOMOR");
      const hasNama = cells.some(
        (c) => c === "NAMA" || c === "NAMA ANGGOTA" || c === "NAMA PENGURUS"
      );
      if (hasNo && hasNama) return i;
    }
    return -1;
  };

  const headerRowIndex = findHeaderRowIndex();
  if (headerRowIndex < 0) {
    return xlsx.utils.sheet_to_json(sh, { defval: "" });
  }

  const header = aoa[headerRowIndex].map(norm);
  const dataRows = aoa.slice(headerRowIndex + 1);

  const colIndex = (candidates) => {
    const candNorm = candidates.map((x) => norm(x));
    for (let i = 0; i < header.length; i++) {
      if (candNorm.includes(header[i])) return i;
    }
    return -1;
  };

  const idxNo = colIndex(["NO", "NOMOR"]);
  const idxNama = colIndex(["NAMA", "NAMA ANGGOTA", "NAMA PENGURUS"]);
  const idxJabat = colIndex([
    "JABATAN/BIDANG",
    "JABATAN BIDANG",
    "JABATAN",
    "BIDANG",
    "JABAT",
    "JAB",
  ]);

  const out = [];
  for (let r = 0; r < dataRows.length; r++) {
    const row = dataRows[r] || [];
    const no = idxNo >= 0 ? cleanText(row[idxNo]) : "";
    const nama = idxNama >= 0 ? cleanText(row[idxNama]) : "";
    const jabatanBidang = idxJabat >= 0 ? cleanText(row[idxJabat]) : "";

    if (!no && !nama && !jabatanBidang) continue;
    out.push({ no, nama, jabatanBidang });
  }

  return out;
}

/**
 * Select2 penduduk (#kelompok_penduduk) round-based
 */
async function pickPendudukByName(page, namaRaw) {
  const base = cleanText(namaRaw);
  const core = coreName(base);

  if (!base) throw new Error("Nama anggota kosong");
  if (!core) throw new Error("Nama anggota tidak valid");

  const queries = buildQueriesNama(base);
  const coreRegex = new RegExp(escapeRegExp(core), "i");

  const select2Clickable = page
    .locator(
      [
        "#select2-kelompok_penduduk-container",
        'span.select2-selection--single[aria-labelledby="select2-kelompok_penduduk-container"]',
        "span.select2-container:has(#select2-kelompok_penduduk-container)",
      ].join(",")
    )
    .first();

  await select2Clickable.waitFor({ state: "visible", timeout: 15000 });

  const MAX_ROUNDS = 3;
  const TYPE_DELAY = 160;
  const STABILIZE_WAIT = 1100;

  for (let round = 1; round <= MAX_ROUNDS; round++) {
    for (const q of queries) {
      const openCount = await page.locator(".select2-container--open").count();
      if (!openCount) await select2Clickable.click();

      const searchField = page.locator("input.select2-search__field").first();
      await searchField.waitFor({ state: "visible", timeout: 10000 });

      await searchField.fill("");
      await sleep(150);
      await searchField.type(q, { delay: TYPE_DELAY });

      await Promise.race([
        page
          .waitForResponse(
            (res) =>
              res.ok() &&
              /select2|penduduk|autocomplete|ajax|cari/i.test(res.url()) &&
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

      const noResult = page.locator(".select2-results__message").first();
      if (await noResult.isVisible().catch(() => false)) continue;

      const option = page
        .locator(".select2-results__option", { hasText: coreRegex })
        .filter({ hasNot: page.locator(".select2-results__message") })
        .first();

      if ((await option.count()) === 0) continue;

      await option.click();

      await page
        .locator(".select2-container--open")
        .waitFor({ state: "detached", timeout: 6000 })
        .catch(() => {});

      const selectedText = cleanText(
        await page.locator("#select2-kelompok_penduduk-container").textContent()
      );

      if (!coreRegex.test(selectedText || "")) continue;

      return;
    }

    await sleep(400);
  }

  throw new Error(`Nama anggota tidak ditemukan: ${base}`);
}

/**
 * Mapping jabatan/bidang excel -> dropdown jabatan
 */
function mapJabatanLabel(jabatanBidangRaw) {
  const j = upper(jabatanBidangRaw);
  if (j.includes("WAKIL")) return "WAKIL KETUA";
  if (j.includes("SEKRET")) return "SEKRETARIS";
  if (j.includes("BENDA")) return "BENDAHARA";
  if (j.includes("KETUA")) return "KETUA";
  return "ANGGOTA";
}

async function writeReportExcel(originalRows, errorMessages, outPath) {
  const headers = [...Object.keys(originalRows[0] || {}), "PESAN_GAGAL"];
  // kalau ada variasi keys antar baris:
  const allKeys = Array.from(
    new Set(originalRows.flatMap((r) => Object.keys(r)))
  );
  const finalHeaders = [...allKeys, "PESAN_GAGAL"];

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("REPORT");

  ws.addRow(finalHeaders);
  ws.getRow(1).font = { bold: true };

  const yellowFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF00" },
  };

  for (let i = 0; i < originalRows.length; i++) {
    const base = originalRows[i];
    const pesan = cleanText(errorMessages[i] || "");
    const rowValues = finalHeaders.map((h) =>
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

// ====== PROSES 1 ROW (dipakai pass 1 & reprocess) ======
async function processOneRow(page, context, row) {
  const no = cleanText(row.no);
  const nama = cleanText(row.nama);
  const jabatanBidang = cleanText(row.jabatanBidang);

  if (!no) return "Kolom no kosong";
  if (!nama) return "Kolom nama kosong";

  await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });

  if (page.url().includes("/siteman")) {
    await ensureLoggedIn(page);
    await context.storageState({ path: STORAGE_STATE_PATH });
    await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });
  }

  await pickPendudukByName(page, nama);

  await page
    .locator("#no_anggota")
    .waitFor({ state: "visible", timeout: 15000 });
  await page.fill("#no_anggota", no);

  const jabatanLabel = mapJabatanLabel(jabatanBidang);
  await page.locator("#jabatan").waitFor({ state: "attached", timeout: 15000 });
  const selJabatan = await page.selectOption("#jabatan", {
    label: jabatanLabel,
  });
  if (!selJabatan || selJabatan.length === 0) {
    return `Jabatan "${jabatanLabel}" tidak tersedia di dropdown`;
  }
  await page.dispatchEvent("#jabatan", "change").catch(() => {});

  await page
    .locator("#no_sk_jabatan")
    .waitFor({ state: "visible", timeout: 15000 });
  await page.fill("#no_sk_jabatan", "-");

  await page
    .locator('input[name="nmr_sk_pengangkatan"]')
    .waitFor({ state: "visible", timeout: 15000 });
  await page.fill('input[name="nmr_sk_pengangkatan"]', SK_PENGANGKATAN_NO);

  await page
    .locator('input[name="tgl_sk_pengangkatan"]')
    .waitFor({ state: "visible", timeout: 15000 });
  await page.fill('input[name="tgl_sk_pengangkatan"]', SK_PENGANGKATAN_TGL);

  await page
    .locator('input[name="nmr_sk_pemberhentian"]')
    .waitFor({ state: "visible", timeout: 15000 });
  await page.fill('input[name="nmr_sk_pemberhentian"]', SK_PEMBERHENTIAN_NO);

  await page
    .locator('input[name="tgl_sk_pemberhentian"]')
    .waitFor({ state: "visible", timeout: 15000 });
  await page.fill('input[name="tgl_sk_pemberhentian"]', SK_PEMBERHENTIAN_TGL);

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

// ====== REPROCESS failed rows max 3x ======
async function reprocessFailed(
  page,
  context,
  rows,
  errorMessages,
  maxRetry = 3
) {
  for (let attempt = 1; attempt <= maxRetry; attempt++) {
    const failedIndexes = [];
    for (let i = 0; i < errorMessages.length; i++) {
      if (cleanText(errorMessages[i])) failedIndexes.push(i);
    }

    if (failedIndexes.length === 0) return;

    console.log(
      `\n=== REPROCESS attempt ${attempt}/${maxRetry} | gagal: ${failedIndexes.length} baris ===`
    );

    // jeda kecil antar attempt
    await sleep(800);

    for (const idx of failedIndexes) {
      const row = rows[idx];
      const no = cleanText(row.no);
      const nama = cleanText(row.nama);

      console.log(
        `  [RETRY ${attempt}] idx=${idx + 1} No=${no} | Nama=${nama}`
      );

      try {
        const err = await processOneRow(page, context, row);
        if (!err) {
          errorMessages[idx] = "";
          console.log("    - Sukses (reprocess)");
        } else {
          errorMessages[idx] = err;
          console.log("    - Masih gagal:", err);
        }
      } catch (e) {
        errorMessages[idx] = cleanText(e.message);
        console.log("    - Exception:", errorMessages[idx]);
      }

      // biar tidak terlalu ngegas ke server
      await sleep(300);
    }
  }
}

async function main() {
  const rows = readKarangTarunaData(EXCEL_PATH);
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
    const row = rows[idx];
    const no = cleanText(row.no);
    const nama = cleanText(row.nama);
    const jabatanBidang = cleanText(row.jabatanBidang);

    console.log(
      `\n[${idx + 1}/${
        rows.length
      }] No=${no} | Nama=${nama} | Jabatan/Bidang=${jabatanBidang}`
    );

    try {
      const err = await processOneRow(page, context, row);
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
