/**
 * input-lpm.js
 * Otomasi input anggota LPM (lembaga_anggota) dari Excel: data_pengurus_lpm.xlsx
 *
 * URL form: http://opensid.local/lembaga_anggota/form/51
 * Field yang diisi:
 * - Nama Anggota (select2 #kelompok_penduduk)   <- kolom "nama"
 * - Nomor Anggota (#no_anggota)                 <- kolom "no"
 * - Jabatan (#jabatan)                          <- kolom "jabat"
 * - Nomor SK Jabatan (#no_sk_jabatan)           <- "-" (biar aman)
 * - Nomor/Tgl SK Pengangkatan                   <- 141/16/DB/01/2023 / 23-02-2023
 * - Nomor/Tgl SK Pemberhentian                  <- "-" / 10-01-2026
 *
 * Output report: report-lpm.xlsx (baris gagal kuning + kolom PESAN_GAGAL)
 */

const { chromium } = require("playwright");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const BASE = "http://opensid.local";
const LOGIN_URL = `${BASE}/siteman`;
const FORM_URL = `${BASE}/lembaga_anggota/form/51`;

const EXCEL_PATH = path.resolve(__dirname, "data_pengurus_lpm.xlsx");
const STORAGE_STATE_PATH = path.resolve(__dirname, "storageState.json");
const REPORT_PATH = path.resolve(__dirname, "report-lpm.xlsx");

const USERNAME = process.env.OPENSID_USER || "admin";
const PASSWORD = process.env.OPENSID_PASS || "Admin123.";

// Konstanta sesuai permintaan
const SK_PENGANGKATAN_NO = "141/16/DB/01/2023";
const SK_PENGANGKATAN_TGL = "23-02-2023";
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

// Nama inti untuk prefix "R." / "Raden" (minimal & cepat)
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
 * Baca Excel dengan header yang mungkin variasi:
 * - no / No / NO
 * - nama / Nama / NAMA
 * - jabat / jabatan / JABATAN / jab
 *
 * Jika header tidak di baris pertama, script akan cari baris yang mengandung minimal: NO + NAMA.
 */
function readPengurusLpmData(filePath) {
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
        (c) => c === "NAMA" || c === "NAMA PENGURUS" || c === "NAMA ANGGOTA"
      );
      if (hasNo && hasNama) return i;
    }
    return -1;
  };

  const headerRowIndex = findHeaderRowIndex();
  if (headerRowIndex < 0) {
    // fallback: anggap baris 1 header
    const rows = xlsx.utils.sheet_to_json(sh, { defval: "" });
    return rows;
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
  const idxJabat = colIndex(["JABAT", "JAB", "JABATAN"]);

  const out = [];
  for (let r = 0; r < dataRows.length; r++) {
    const row = dataRows[r] || [];
    const no = idxNo >= 0 ? cleanText(row[idxNo]) : "";
    const nama = idxNama >= 0 ? cleanText(row[idxNama]) : "";
    const jabat = idxJabat >= 0 ? cleanText(row[idxJabat]) : "";

    if (!no && !nama && !jabat) continue;

    out.push({
      no,
      nama,
      jabat,
    });
  }

  return out;
}

/**
 * Pilih penduduk pada select2 #kelompok_penduduk (Nama Anggota)
 * Model round-based: 1 round mencoba semua query (asli -> Raden -> inti), baru ulang round.
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
 * Mapping jabatan dari Excel ke label dropdown.
 * (minimal, cepat, tahan variasi)
 */
function mapJabatanLabel(jabatRaw) {
  const j = upper(jabatRaw);

  if (j.includes("WAKIL")) return "WAKIL KETUA";
  if (j.includes("SEKRET")) return "SEKRETARIS";
  if (j.includes("BENDA")) return "BENDAHARA";
  if (j.includes("KETUA")) return "KETUA";
  return "ANGGOTA";
}

async function writeReportExcel(originalRows, errorMessages, outPath) {
  const originalKeys = Array.from(
    new Set(originalRows.flatMap((r) => Object.keys(r)))
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

  for (let i = 0; i < originalRows.length; i++) {
    const base = originalRows[i];
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

async function main() {
  const rows = readPengurusLpmData(EXCEL_PATH);
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

  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx];

    const no = cleanText(row.no);
    const nama = cleanText(row.nama);
    const jabat = cleanText(row.jabat);

    console.log(
      `\n[${idx + 1}/${rows.length}] No=${no} | Nama=${nama} | Jabat=${jabat}`
    );

    if (!no) {
      errorMessages[idx] = "Kolom no kosong";
      console.log("  - Gagal:", errorMessages[idx]);
      continue;
    }
    if (!nama) {
      errorMessages[idx] = "Kolom nama kosong";
      console.log("  - Gagal:", errorMessages[idx]);
      continue;
    }

    try {
      await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });

      // kalau kelempar login lagi
      if (page.url().includes("/siteman")) {
        await ensureLoggedIn(page);
        await context.storageState({ path: STORAGE_STATE_PATH });
        await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });
      }

      // pilih nama anggota
      await pickPendudukByName(page, nama);

      // isi nomor anggota (dari kolom no)
      await page
        .locator("#no_anggota")
        .waitFor({ state: "visible", timeout: 15000 });
      await page.fill("#no_anggota", no);

      // pilih jabatan
      const jabatanLabel = mapJabatanLabel(jabat);
      await page
        .locator("#jabatan")
        .waitFor({ state: "attached", timeout: 15000 });
      const selJabatan = await page.selectOption("#jabatan", {
        label: jabatanLabel,
      });
      if (!selJabatan || selJabatan.length === 0) {
        throw new Error(`Jabatan "${jabatanLabel}" tidak tersedia di dropdown`);
      }
      await page.dispatchEvent("#jabatan", "change").catch(() => {});

      // nomor SK jabatan -> "-"
      await page
        .locator("#no_sk_jabatan")
        .waitFor({ state: "visible", timeout: 15000 });
      await page.fill("#no_sk_jabatan", "-");

      // SK pengangkatan
      await page
        .locator('input[name="nmr_sk_pengangkatan"]')
        .waitFor({ state: "visible", timeout: 15000 });
      await page.fill('input[name="nmr_sk_pengangkatan"]', SK_PENGANGKATAN_NO);

      await page
        .locator('input[name="tgl_sk_pengangkatan"]')
        .waitFor({ state: "visible", timeout: 15000 });
      await page.fill('input[name="tgl_sk_pengangkatan"]', SK_PENGANGKATAN_TGL);

      // SK pemberhentian
      await page
        .locator('input[name="nmr_sk_pemberhentian"]')
        .waitFor({ state: "visible", timeout: 15000 });
      await page.fill(
        'input[name="nmr_sk_pemberhentian"]',
        SK_PEMBERHENTIAN_NO
      );

      await page
        .locator('input[name="tgl_sk_pemberhentian"]')
        .waitFor({ state: "visible", timeout: 15000 });
      await page.fill(
        'input[name="tgl_sk_pemberhentian"]',
        SK_PEMBERHENTIAN_TGL
      );

      // optional: periode & keterangan dikosongkan (tidak diminta)
      // await page.fill('input[name="periode"]', "");
      // await page.fill('textarea[name="keterangan"]', "");

      // simpan
      const simpanBtn = page.locator('button[type="submit"]', {
        hasText: /Simpan/i,
      });
      await simpanBtn.waitFor({ state: "visible", timeout: 15000 });

      await Promise.allSettled([
        page.waitForLoadState("networkidle", { timeout: 15000 }),
        simpanBtn.click(),
      ]);

      errorMessages[idx] = "";
      console.log("  - Sukses");
    } catch (err) {
      errorMessages[idx] = cleanText(err.message);
      console.log("  - Gagal:", errorMessages[idx]);
    }
  }

  await browser.close();

  await writeReportExcel(rows, errorMessages, REPORT_PATH);
  console.log(`\nSelesai. Report dibuat: ${REPORT_PATH}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
