const { chromium } = require("playwright");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const BASE = "http://opensid.local";
const LOGIN_URL = `${BASE}/siteman`;
const FORM_URL = `${BASE}/peserta_bantuan/form/3`;

const EXCEL_PATH = path.resolve(__dirname, "kpm.xlsx");
const STORAGE_STATE_PATH = path.resolve(__dirname, "storageState.json");
const REPORT_PATH = path.resolve(__dirname, "report-kpm.xlsx");

const USERNAME = process.env.OPENSID_USER || "admin";
const PASSWORD = process.env.OPENSID_PASS || "Admin123.";

function normalizeNik(value) {
  return String(value ?? "").replace(/\D/g, "");
}
function cleanText(s) {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();
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

async function openSelect2AndPickByNameValidateNik(
  page,
  namaKpm,
  expectedNikRaw
) {
  const expectedNik = normalizeNik(expectedNikRaw);
  if (!expectedNik) throw new Error("NIK kosong/tidak valid");

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

  async function trySearchAndPick(query) {
    await searchField.fill(query);

    await page.waitForSelector(".select2-results__option", { timeout: 10000 });

    const nikOption = page
      .locator(".select2-results__option", { hasText: expectedNik })
      .filter({ hasNot: page.locator(".select2-results__message") })
      .first();

    if ((await nikOption.count()) > 0) {
      await nikOption.click();
      return true;
    }
    return false;
  }

  // coba nama dulu, kalau gagal fallback NIK
  let picked = await trySearchAndPick(namaKpm);
  if (!picked) picked = await trySearchAndPick(expectedNik);

  if (!picked) {
    throw new Error(`Tidak ditemukan data dengan NIK ${expectedNik}`);
  }

  const selectedText = cleanText(
    await page.locator("#select2-nik_bantuan-container").textContent()
  );
  const selectedNik = normalizeNik(selectedText);

  if (!selectedNik.includes(expectedNik)) {
    throw new Error(`NIK tidak cocok. Terpilih: "${selectedText}"`);
  }
}

async function writeReportExcel(originalRows, errorMessages, outPath) {
  // header = kolom asli + 1 kolom pesan gagal
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

    const rowValues = headers.map((h) => {
      if (h === "PESAN_GAGAL") return pesan;
      return base[h] ?? "";
    });

    const row = ws.addRow(rowValues);

    // Marking kuning kalau ada pesan gagal
    if (pesan) {
      row.eachCell((cell) => {
        cell.fill = yellowFill;
      });
    }
  }

  // auto width sederhana
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

  const COL_NAMA = "NAMA KPM";
  const COL_NIK = "NIK";

  // 1:1 dengan rows — kosong berarti sukses
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
    const namaKpm = cleanText(row[COL_NAMA]);
    const nik = normalizeNik(row[COL_NIK]);

    console.log(`\n[${idx + 1}/${rows.length}] ${namaKpm} | ${nik}`);

    if (!namaKpm || !nik) {
      errorMessages[idx] = "Nama KPM atau NIK kosong";
      console.log("  - Gagal: Nama/NIK kosong");
      continue;
    }

    try {
      await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });

      if (page.url().includes("/siteman")) {
        await ensureLoggedIn(page);
        await context.storageState({ path: STORAGE_STATE_PATH });
        await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });
      }

      await openSelect2AndPickByNameValidateNik(page, namaKpm, nik);

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

      // sukses = pesan kosong
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
