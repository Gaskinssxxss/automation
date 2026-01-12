const { chromium } = require("playwright");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const BASE = "https://dev-bayan.ntbprov.go.id/index.php";
const LOGIN_URL = `${BASE}/siteman`;
const FORM_URL = `${BASE}/peserta_bantuan/form/1042002`;

const EXCEL_PATH = path.resolve(__dirname, "pbp.xlsx");
const STORAGE_STATE_PATH = path.resolve(__dirname, "storageState.json");
const REPORT_PATH = path.resolve(__dirname, "report-pbp.xlsx");

const USERNAME = process.env.OPENSID_USER || "administrator";
const PASSWORD = process.env.OPENSID_PASS || "@Muhammad270901";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function cleanText(s) {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();
}
function normalizeName(s) {
  return cleanText(s).toUpperCase();
}
function escapeRegExp(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Ambil keyword dusun dari alamat.
 * Contoh yang didukung:
 * - "DUSUN KARANGANYAR RT 01 RW 02"
 * - "Dusun: Karanganyar, RT 1 RW 2"
 * - "Kp. Karanganyar" (fallback)
 *
 * Output: "KARANGANYAR" (tanpa RT/RW)
 * Kalau gagal, return "" (nanti validasi cuma nama).
 */
function extractDusunKeyword(alamatRaw) {
  const alamat = normalizeName(alamatRaw);

  if (!alamat) return "";

  // Pola umum: DUSUN <nama>
  // Tangkap sampai sebelum RT/RW/NO/JLN/DS/KEC/KAB/dll atau koma
  const m1 = alamat.match(
    /\bDUSUN\b\s*[:\-]?\s*([A-Z0-9'./\-\s]{3,50}?)(?=\bRT\b|\bRW\b|,|\bDESA\b|\bKEL\b|\bKEC\b|\bKAB\b|\bKOTA\b|\bNO\b|\bJL\b|\bJLN\b|$)/
  );
  if (m1 && m1[1]) return cleanText(m1[1]).toUpperCase();

  // Pola alternatif: "KP" / "KAMPUNG" / "KMP"
  const m2 = alamat.match(
    /\b(KP|KAMPUNG|KMP)\b\.?\s*[:\-]?\s*([A-Z0-9'./\-\s]{3,50}?)(?=\bRT\b|\bRW\b|,|\bDESA\b|\bKEL\b|\bKEC\b|\bKAB\b|\bKOTA\b|\bNO\b|\bJL\b|\bJLN\b|$)/
  );
  if (m2 && m2[2]) return cleanText(m2[2]).toUpperCase();

  // Kalau tidak ada pattern, coba ambil token pertama yang bukan RT/RW/angka pendek
  // (fallback lemah tapi lumayan)
  const tokens = alamat
    .replace(/[^\w\s]/g, " ")
    .split(/\s+/)
    .filter(Boolean)
    .filter(
      (t) =>
        ![
          "RT",
          "RW",
          "NO",
          "JL",
          "JLN",
          "DESA",
          "KEL",
          "KEC",
          "KAB",
          "KOTA",
        ].includes(t)
    )
    .filter((t) => t.length >= 4 && !/^\d+$/.test(t));

  return tokens[0] ? tokens[0].toUpperCase() : "";
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
 * Select2 penduduk (lebih sabar & lebih valid):
 * - search pakai NAMA PBP
 * - pilih opsi yang mengandung NAMA + (DUSUN jika ada)
 * - validasi selection mengandung NAMA + (DUSUN jika ada)
 */
async function pickPendudukByNameAndDusun(page, namaPbpRaw, alamatPbpRaw) {
  const namaPbp = cleanText(namaPbpRaw);
  if (!namaPbp) throw new Error("Nama PBP kosong");

  const dusunKey = extractDusunKeyword(alamatPbpRaw); // bisa ""
  const nameRegex = new RegExp(escapeRegExp(namaPbp), "i");
  const dusunRegex = dusunKey ? new RegExp(escapeRegExp(dusunKey), "i") : null;

  const select2Clickable = page
    .locator(
      [
        "#select2-nik_bantuan-container",
        'span.select2-selection--single[aria-labelledby="select2-nik_bantuan-container"]',
        "span.select2-container:has(#select2-nik_bantuan-container)",
      ].join(",")
    )
    .first();

  await select2Clickable.waitFor({ state: "visible", timeout: 15000 });
  await select2Clickable.click();

  const searchField = page.locator("input.select2-search__field").first();
  await searchField.waitFor({ state: "visible", timeout: 10000 });

  const MAX_RETRY = 3;

  for (let attempt = 1; attempt <= MAX_RETRY; attempt++) {
    // ketik pelan
    await searchField.fill("");
    await sleep(200);
    await searchField.type(namaPbp, { delay: 200 }); // lebih lambat biar stabil

    // tunggu ajax/delay stabilisasi
    await Promise.race([
      page
        .waitForResponse(
          (res) =>
            res.ok() &&
            /select2|penduduk|autocomplete|ajax|cari/i.test(res.url()) &&
            (res.request().method() === "GET" ||
              res.request().method() === "POST"),
          { timeout: 6000 }
        )
        .catch(() => null),
      sleep(900),
    ]);

    await page.waitForSelector(".select2-results__option", { timeout: 10000 });

    const noResult = page.locator(".select2-results__message").first();
    if (await noResult.isVisible().catch(() => false)) {
      await sleep(800 + attempt * 250);
      continue;
    }

    // 1) Prefer: opsi yang mengandung NAMA + DUSUN (kalau dusunKey ada)
    let option = null;

    if (dusunRegex) {
      option = page
        .locator(".select2-results__option", { hasText: nameRegex })
        .filter({ hasText: dusunRegex })
        .filter({ hasNot: page.locator(".select2-results__message") })
        .first();

      if ((await option.count()) === 0) option = null;
    }

    // 2) Fallback: opsi yang mengandung NAMA saja
    if (!option) {
      const optByName = page
        .locator(".select2-results__option", { hasText: nameRegex })
        .filter({ hasNot: page.locator(".select2-results__message") })
        .first();

      if ((await optByName.count()) === 0) {
        await sleep(800 + attempt * 250);
        continue;
      }
      option = optByName;
    }

    await option.click();

    // tunggu dropdown menutup
    await page
      .locator(".select2-container--open")
      .waitFor({ state: "detached", timeout: 5000 })
      .catch(() => {});

    const selectedText = cleanText(
      await page.locator("#select2-nik_bantuan-container").textContent()
    );
    const selectedUpper = normalizeName(selectedText);

    // Validasi nama
    if (!selectedUpper.includes(normalizeName(namaPbp))) {
      await sleep(800 + attempt * 250);
      await select2Clickable.click();
      await searchField.waitFor({ state: "visible", timeout: 10000 });
      continue;
    }

    // Validasi dusun kalau ada keyword
    if (dusunKey && !selectedUpper.includes(dusunKey)) {
      // kalau dusun tidak cocok, coba lagi
      await sleep(800 + attempt * 250);
      await select2Clickable.click();
      await searchField.waitFor({ state: "visible", timeout: 10000 });
      continue;
    }

    return; // sukses
  }

  if (dusunKey) {
    throw new Error(
      `Tidak ditemukan hasil valid untuk "${namaPbp}" dengan dusun "${dusunKey}"`
    );
  }
  throw new Error(`Tidak ditemukan hasil valid untuk nama: ${namaPbp}`);
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
    fgColor: { argb: "FFFFFF00" }, // kuning
  };

  for (let i = 0; i < originalRows.length; i++) {
    const base = originalRows[i];
    const pesan = cleanText(errorMessages[i] || "");

    const rowValues = headers.map((h) => {
      if (h === "PESAN_GAGAL") return pesan;
      return base[h] ?? "";
    });

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
  const wb = xlsx.readFile(EXCEL_PATH);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

  const COL_NAMA = "NAMA PBP";
  const COL_ALAMAT = "ALAMAT PBP";

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
    const namaPbp = cleanText(row[COL_NAMA]);
    const alamatPbp = cleanText(row[COL_ALAMAT]);

    console.log(`\n[${idx + 1}/${rows.length}] ${namaPbp}`);

    if (!namaPbp) {
      errorMessages[idx] = "Nama PBP kosong";
      console.log("  - Gagal: Nama PBP kosong");
      continue;
    }

    try {
      await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });

      if (page.url().includes("/siteman")) {
        await ensureLoggedIn(page);
        await context.storageState({ path: STORAGE_STATE_PATH });
        await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });
      }

      await pickPendudukByNameAndDusun(page, namaPbp, alamatPbp);

      await page
        .locator("#no_id_kartu")
        .waitFor({ state: "visible", timeout: 15000 });
      await page.fill("#no_id_kartu", "-");

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
