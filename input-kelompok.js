const { chromium } = require("playwright");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const BASE = "http://opensid.local";
const LOGIN_URL = `${BASE}/siteman`;
const FORM_URL = `${BASE}/kelompok/form`;

const EXCEL_PATH = path.resolve(__dirname, "kelompok.xlsx");
const STORAGE_STATE_PATH = path.resolve(__dirname, "storageState.json");
const REPORT_PATH = path.resolve(__dirname, "report-kelompok.xlsx");

const USERNAME = process.env.OPENSID_USER || "admin";
const PASSWORD = process.env.OPENSID_PASS || "Admin123.";
const KODE_START = Number(process.env.KODE_START || "1");

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

// ====== Nama inti untuk kasus "R." / "Raden" ======
function coreName(nama) {
  let s = cleanText(nama);
  s = s.replace(/^\s*R\s*\.\s*/i, ""); // "R. John" -> "John"
  s = s.replace(/^\s*RADEN\s+/i, ""); // "Raden John" -> "John"
  return cleanText(s);
}
function buildKetuaQueriesFast(namaKetuaRaw) {
  const base = cleanText(namaKetuaRaw);
  const core = coreName(base);

  const queries = [];
  if (base) queries.push(base);

  // hanya tambah varian kalau memang ada prefix
  if (core && core !== base) {
    queries.push(`Raden ${core}`);
    queries.push(core);
  }

  return [...new Set(queries)];
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
 * Baca kelompok.xlsx (format header 2 baris):
 * - cari baris yang memuat "Nama Poktan"
 * - header baris berikutnya biasanya memuat "Ketua/Sekretaris/Bendahara"
 */
function readKelompokData(filePath) {
  const wb = xlsx.readFile(filePath);
  const sheetNames = wb.SheetNames;

  for (const sn of sheetNames) {
    const sh = wb.Sheets[sn];
    const aoa = xlsx.utils.sheet_to_json(sh, { header: 1, defval: "" });

    const header1Index = aoa.findIndex((row) =>
      row.some((cell) => upper(cell) === "NAMA POKTAN")
    );
    if (header1Index < 0) continue;

    const header2Index = header1Index + 1;
    const h1 = aoa[header1Index] || [];
    const h2 = aoa[header2Index] || [];

    const findCol = (h1Text, h2Text = null) => {
      for (let c = 0; c < Math.max(h1.length, h2.length); c++) {
        const a = upper(h1[c]);
        const b = upper(h2[c]);
        if (a === upper(h1Text) && (h2Text ? b === upper(h2Text) : true))
          return c;
      }
      return -1;
    };

    const idxNama = findCol("Nama Poktan");
    const idxAlamat = findCol("Alamat");
    const idxTahun = findCol("Tahun Berdiri");

    let idxKetua = findCol("Nama Pengurus", "Ketua");
    if (idxKetua < 0) idxKetua = findCol("Ketua");

    if (idxNama < 0) continue;

    const dataStart = header2Index + 1;
    const out = [];

    for (let r = dataStart; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const namaPoktan = cleanText(row[idxNama]);
      const alamat = idxAlamat >= 0 ? cleanText(row[idxAlamat]) : "";
      const tahun = idxTahun >= 0 ? cleanText(row[idxTahun]) : "";
      const namaKetua = idxKetua >= 0 ? cleanText(row[idxKetua]) : "";

      if (!namaPoktan && !alamat && !tahun && !namaKetua) continue;

      out.push({
        "Nama Poktan": namaPoktan,
        Alamat: alamat,
        "Tahun Berdiri": tahun,
        "Nama Ketua": namaKetua,
        __SHEET: sn,
        __ROW: r + 1,
      });
    }

    if (out.length > 0) {
      console.log(
        `Excel terbaca dari sheet: "${sn}" (baris data: ${out.length})`
      );
      return out;
    }
  }

  throw new Error(`Tidak menemukan header "Nama Poktan" di semua sheet.`);
}

/**
 * Select2 Ketua: ROUND-based retry (cepat ganti kondisi)
 * - dalam 1 round: coba semua query (asli -> Raden -> inti) masing2 1x
 * - kalau no result: langsung lanjut query berikutnya
 * - ulang round beberapa kali kalau server lambat
 * - opsi dipilih & divalidasi pakai nama inti (core)
 */
async function pickKetuaByName(page, namaKetuaRaw) {
  const base = cleanText(namaKetuaRaw);
  const core = coreName(base);

  if (!base) throw new Error("Nama Ketua kosong");
  if (!core) throw new Error("Nama Ketua tidak valid");

  const queries = buildKetuaQueriesFast(base);
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

  const MAX_ROUNDS = 3; // total putaran
  const TYPE_DELAY = 160;
  const STABILIZE_WAIT = 1100;

  for (let round = 1; round <= MAX_ROUNDS; round++) {
    for (const q of queries) {
      // buka dropdown kalau belum open
      const openCount = await page.locator(".select2-container--open").count();
      if (!openCount) await select2Clickable.click();

      const searchField = page.locator("input.select2-search__field").first();
      await searchField.waitFor({ state: "visible", timeout: 10000 });

      await searchField.fill("");
      await sleep(150);
      await searchField.type(q, { delay: TYPE_DELAY });

      // tunggu ajax / stabilisasi
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

      // kalau no results, langsung lanjut query berikutnya
      const noResult = page.locator(".select2-results__message").first();
      if (await noResult.isVisible().catch(() => false)) {
        continue;
      }

      // pilih opsi berdasar NAMA INTI (core)
      const option = page
        .locator(".select2-results__option", { hasText: coreRegex })
        .filter({ hasNot: page.locator(".select2-results__message") })
        .first();

      if ((await option.count()) === 0) {
        continue;
      }

      await option.click();

      // tunggu dropdown menutup
      await page
        .locator(".select2-container--open")
        .waitFor({ state: "detached", timeout: 6000 })
        .catch(() => {});

      // validasi juga berdasarkan nama inti
      const selectedText = cleanText(
        await page.locator("#select2-kelompok_penduduk-container").textContent()
      );

      if (!coreRegex.test(selectedText || "")) {
        continue;
      }

      return; // sukses
    }

    // jeda kecil sebelum round berikutnya
    await sleep(400);
  }

  throw new Error(`Ketua tidak ditemukan: ${base}`);
}

async function writeReportExcel(originalRows, errorMessages, outPath) {
  const cleaned = originalRows.map((r) => {
    const copy = { ...r };
    delete copy.__SHEET;
    delete copy.__ROW;
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

async function main() {
  const rows = readKelompokData(EXCEL_PATH);
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

  let kodeCounter = KODE_START;

  for (let idx = 0; idx < rows.length; idx++) {
    const r = rows[idx];

    const namaPoktan = cleanText(r["Nama Poktan"]);
    const namaKetua = cleanText(r["Nama Ketua"]);
    const alamat = cleanText(r["Alamat"]);
    const tahun = cleanText(r["Tahun Berdiri"]);

    console.log(`\n[${idx + 1}/${rows.length}] Kelompok: ${namaPoktan}`);

    if (!namaPoktan) {
      errorMessages[idx] = "Nama Poktan kosong";
      console.log("  - Gagal: Nama Poktan kosong");
      continue;
    }
    if (!namaKetua) {
      errorMessages[idx] = "Nama Ketua kosong";
      console.log("  - Gagal: Nama Ketua kosong");
      continue;
    }

    const kodeKelompok = String(kodeCounter);
    kodeCounter += 1;

    try {
      await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });

      if (page.url().includes("/siteman")) {
        await ensureLoggedIn(page);
        await context.storageState({ path: STORAGE_STATE_PATH });
        await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });
      }

      await page.locator("#nama").waitFor({ state: "visible", timeout: 15000 });
      await page.fill("#nama", namaPoktan);

      await page.locator("#kode").waitFor({ state: "visible", timeout: 15000 });
      await page.fill("#kode", kodeKelompok);

      await page
        .locator("#no_sk_pendirian")
        .waitFor({ state: "visible", timeout: 15000 });
      await page.fill("#no_sk_pendirian", "-");

      await page
        .locator("#id_master")
        .waitFor({ state: "attached", timeout: 15000 });
      const selected = await page.selectOption("#id_master", {
        label: "Gabungan Kelompok Tani (GAPOKTAN)",
      });
      if (!selected || selected.length === 0) {
        throw new Error(
          'Kategori "Gabungan Kelompok Tani (GAPOKTAN)" tidak tersedia'
        );
      }
      await page.dispatchEvent("#id_master", "change").catch(() => {});

      await pickKetuaByName(page, namaKetua);

      const deskripsi = `Alamat : ${alamat || "-"}\nTahun Berdiri : ${
        tahun || "-"
      }`;
      await page
        .locator('textarea[name="keterangan"]')
        .waitFor({ state: "visible", timeout: 15000 });
      await page.fill('textarea[name="keterangan"]', deskripsi);

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
