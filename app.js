/* CMS Contract Generator — FULL app.js (DELETE & REPLACE)
   - PIN gate: 0623
   - Upload CSV + Buyer DOCX template + Seller DOCX template EACH SESSION (templates not stored on web)
   - For each CSV row: generate Buyer + Seller editable DOCX using docxtemplater with {{ }} delimiters
   - ZIP output:
       /Buyer Contracts/{Contract#}-{Buyer}.docx
       /Seller Contracts/{Consignor}-{Contract#}.docx
   - Fixes "stuck at Reading CSV..." by parsing CSV directly from File via Papa.parse(file, ...)
   - Matches IDs in your latest post-auction style index.html:
       exitBtn, pinScreen, appScreen, pinInput, pinBtn, pinMsg,
       dropCsv, dropBuyer, dropSeller,
       csvPicker, buyerPicker, sellerPicker,
       csvPickBtn, buyerPickBtn, sellerPickBtn,
       csvPill, buyerPill, sellerPill,
       genBtn, status
*/

(() => {
  "use strict";

  const CONFIG = {
    PIN: "0623",
    ZIP_FOLDERS: {
      buyer: "Buyer Contracts",
      seller: "Seller Contracts",
    },
    REQUIRED_COLS: [
      "Contract #",
      "Consignor",
      "Buyer",
      "Lot Number #2",
      "Head Count",
    ],
    // Your DOCX templates use {{Tag}} style
    DELIMS: { start: "{{", end: "}}" },
  };

  // ------------------------------------------------------------
  // IMPORTANT: Prevent the browser from "opening" dropped files
  // ------------------------------------------------------------
  window.addEventListener("dragover", (e) => e.preventDefault(), false);
  window.addEventListener("drop", (e) => e.preventDefault(), false);

  // ------------------------------------------------------------
  // DOM
  // ------------------------------------------------------------
  const $ = (id) => document.getElementById(id);

  const pinScreen = $("pinScreen");
  const appScreen = $("appScreen");
  const pinInput = $("pinInput");
  const pinBtn = $("pinBtn");
  const pinMsg = $("pinMsg");

  const exitBtn = $("exitBtn");

  const dropCsv = $("dropCsv");
  const dropBuyer = $("dropBuyer");
  const dropSeller = $("dropSeller");

  const csvPicker = $("csvPicker");
  const buyerPicker = $("buyerPicker");
  const sellerPicker = $("sellerPicker");

  const csvPickBtn = $("csvPickBtn");
  const buyerPickBtn = $("buyerPickBtn");
  const sellerPickBtn = $("sellerPickBtn");

  const csvPill = $("csvPill");
  const buyerPill = $("buyerPill");
  const sellerPill = $("sellerPill");

  const genBtn = $("genBtn");
  const statusEl = $("status");

  // ------------------------------------------------------------
  // Sanity: ensure required DOM exists (helps diagnose “nothing happens”)
  // ------------------------------------------------------------
  function assertDom() {
    const missing = [];
    [
      ["pinScreen", pinScreen],
      ["appScreen", appScreen],
      ["pinInput", pinInput],
      ["pinBtn", pinBtn],
      ["pinMsg", pinMsg],
      ["exitBtn", exitBtn],
      ["dropCsv", dropCsv],
      ["dropBuyer", dropBuyer],
      ["dropSeller", dropSeller],
      ["csvPicker", csvPicker],
      ["buyerPicker", buyerPicker],
      ["sellerPicker", sellerPicker],
      ["csvPickBtn", csvPickBtn],
      ["buyerPickBtn", buyerPickBtn],
      ["sellerPickBtn", sellerPickBtn],
      ["csvPill", csvPill],
      ["buyerPill", buyerPill],
      ["sellerPill", sellerPill],
      ["genBtn", genBtn],
      ["status", statusEl],
    ].forEach(([name, el]) => {
      if (!el) missing.push(name);
    });

    if (missing.length) {
      console.error("Missing DOM IDs:", missing);
      setStatus(
        "ERROR: Missing required HTML element IDs:\n" +
          missing.map((m) => `- ${m}`).join("\n") +
          "\n\nFix: ensure your index.html includes these IDs exactly."
      );
      return false;
    }
    return true;
  }

  // ------------------------------------------------------------
  // Library checks
  // ------------------------------------------------------------
  function getPizZip() {
    return window.PizZip || window.pizzip || window.Pizzip || null;
  }
  function getDocxtemplater() {
    return window.docxtemplater || window.Docxtemplater || window.DocxTemplater || null;
  }
  function requireLibs() {
    const missing = [];
    if (typeof Papa === "undefined") missing.push("PapaParse (Papa) — /lib/papa.min.js");
    if (!getPizZip()) missing.push("PizZip — /lib/pizzip.min.js");
    if (!getDocxtemplater()) missing.push("docxtemplater — /lib/docxtemplater.min.js");
    if (typeof JSZip === "undefined") missing.push("JSZip — /lib/jszip.min.js");
    if (typeof saveAs === "undefined") missing.push("FileSaver (saveAs) — /lib/FileSaver.min.js");

    if (missing.length) {
      console.error("Missing libs:", missing);
      setStatus(
        "ERROR: Missing required libraries:\n" +
          missing.map((m) => `- ${m}`).join("\n") +
          "\n\nFix: confirm these script tags exist in index.html (before app.js) and filenames match exactly."
      );
      return false;
    }
    return true;
  }

  // ------------------------------------------------------------
  // State
  // ------------------------------------------------------------
  let csvRows = [];
  let buyerTemplateBytes = null; // Uint8Array
  let sellerTemplateBytes = null; // Uint8Array

  // ------------------------------------------------------------
  // UI helpers
  // ------------------------------------------------------------
  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  function logLine(msg) {
    statusEl.textContent = `${msg}\n\n${statusEl.textContent}`;
  }

  function setPill(pillEl, state, text) {
    pillEl.classList.remove("ok", "bad", "warn");
    pillEl.classList.add(state);
    pillEl.textContent = text;
  }

  function refreshGenerateButton() {
    genBtn.disabled = !(csvRows.length > 0 && buyerTemplateBytes && sellerTemplateBytes);
  }

  // ------------------------------------------------------------
  // Filename safety
  // ------------------------------------------------------------
  function sanitizeFilePart(s) {
    const str = String(s ?? "").trim();
    if (!str) return "UNKNOWN";
    return str
      .replace(/[\/\\:*?"<>|]/g, "-")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/^[.\s]+|[.\s]+$/g, "");
  }

  function dedupeName(filename, usedSet) {
    const lower = (x) => String(x).toLowerCase();
    const base = filename.replace(/\.docx$/i, "");
    let name = filename;
    let n = 2;
    while (usedSet.has(lower(name))) {
      name = `${base} (${n}).docx`;
      n++;
    }
    usedSet.add(lower(name));
    return name;
  }

  // ------------------------------------------------------------
  // Drag/drop helpers
  // ------------------------------------------------------------
  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }
  function addDragUI(el, on) {
    el.classList.toggle("dragover", !!on);
  }

  function wireDropZone(el, acceptFn) {
    ["dragenter", "dragover", "dragleave", "drop"].forEach((evt) => {
      el.addEventListener(evt, preventDefaults, false);
    });

    ["dragenter", "dragover"].forEach((evt) => {
      el.addEventListener(evt, () => addDragUI(el, true), false);
    });

    ["dragleave", "drop"].forEach((evt) => {
      el.addEventListener(evt, () => addDragUI(el, false), false);
    });

    el.addEventListener("drop", async (e) => {
      const file = e.dataTransfer?.files?.[0];
      if (!file) return;
      setStatus(`Dropped file: ${file.name}`);
      await acceptFn(file);
    });
  }

  // ------------------------------------------------------------
  // PIN gate
  // ------------------------------------------------------------
  function unlock() {
    pinScreen.classList.add("hide");
    appScreen.classList.remove("hide");
    setStatus("Unlocked. Upload CSV + Buyer template + Seller template, then Generate Contracts ZIP.");
  }

  function handlePinSubmit() {
    const val = String(pinInput.value || "").trim();
    if (val === CONFIG.PIN) {
      pinMsg.textContent = "";
      unlock();
    } else {
      pinMsg.textContent = "Incorrect PIN.";
      pinMsg.style.color = "#fecaca";
    }
  }

  // ------------------------------------------------------------
  // Read file as Uint8Array for DOCX templates
  // ------------------------------------------------------------
  function fileToUint8Array(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onerror = () => reject(new Error("Failed to read file."));
      r.onabort = () => reject(new Error("File read aborted."));
      r.onload = () => resolve(new Uint8Array(r.result));
      r.readAsArrayBuffer(file);
    });
  }

  // ------------------------------------------------------------
  // CSV parsing (KEY FIX): Parse directly from File with Papa.parse(file,...)
  // Avoids hangs from file.text() on some browsers/large files.
  // ------------------------------------------------------------
  function parseCsvFileWithPapa(file) {
    return new Promise((resolve, reject) => {
      let settled = false;

      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        dynamicTyping: false,
        // NOTE: worker:true can be flaky on some setups and requires blob workers.
        // Keep it false for maximum compatibility.
        worker: false,
        complete: (results) => {
          settled = true;
          resolve(results);
        },
        error: (err) => {
          settled = true;
          reject(err);
        },
      });

      // Failsafe: if something truly wedges, throw a helpful error instead of “stuck”
      setTimeout(() => {
        if (!settled) {
          reject(
            new Error(
              "CSV parse timed out. This usually means the file is extremely large or the browser is blocking file parsing."
            )
          );
        }
      }, 30000); // 30s failsafe
    });
  }

  async function handleCsvFile(file) {
    try {
      setStatus(`Reading CSV: ${file.name} ...`);
      setPill(csvPill, "warn", "Reading CSV...");

      const parsed = await parseCsvFileWithPapa(file);

      // surface parse errors (Papa puts some in results.errors too)
      if (parsed.errors && parsed.errors.length) {
        setPill(csvPill, "bad", "CSV parse error");
        setStatus(
          "CSV parse errors:\n" +
            parsed.errors.slice(0, 25).map((e) => `${e.message} (row ${e.row})`).join("\n") +
            (parsed.errors.length > 25 ? `\n... plus ${parsed.errors.length - 25} more` : "")
        );
        csvRows = [];
        refreshGenerateButton();
        return;
      }

      const headers = parsed.meta?.fields || [];
      const missing = CONFIG.REQUIRED_COLS.filter((c) => !headers.includes(c));

      if (missing.length) {
        setPill(csvPill, "bad", "Missing columns");
        setStatus(
          "CSV is missing required columns:\n" +
            missing.map((m) => `- ${m}`).join("\n") +
            "\n\nHeaders found:\n" +
            headers.join(", ")
        );
        csvRows = [];
        refreshGenerateButton();
        return;
      }

      const rows = (parsed.data || []).filter(
        (r) => r && Object.values(r).some((v) => String(v ?? "").trim() !== "")
      );

      csvRows = rows;

      setPill(csvPill, "ok", `${file.name} (${rows.length} rows)`);
      setStatus(`CSV loaded: ${rows.length} rows.\nNow upload BOTH DOCX templates.`);
      refreshGenerateButton();
    } catch (err) {
      console.error(err);
      setPill(csvPill, "bad", "CSV failed");
      setStatus("CSV read/parse failed:\n" + (err?.message || String(err)));
      csvRows = [];
      refreshGenerateButton();
    }
  }

  // ------------------------------------------------------------
  // Template handlers
  // ------------------------------------------------------------
  async function handleBuyerTemplate(file) {
    try {
      setStatus(`Reading Buyer template: ${file.name} ...`);
      setPill(buyerPill, "warn", "Reading Buyer template...");
      buyerTemplateBytes = await fileToUint8Array(file);
      setPill(buyerPill, "ok", file.name);
      logLine("Buyer template loaded.");
      refreshGenerateButton();
    } catch (err) {
      console.error(err);
      buyerTemplateBytes = null;
      setPill(buyerPill, "bad", "Buyer template failed");
      setStatus("Buyer template read failed:\n" + (err?.message || String(err)));
      refreshGenerateButton();
    }
  }

  async function handleSellerTemplate(file) {
    try {
      setStatus(`Reading Seller template: ${file.name} ...`);
      setPill(sellerPill, "warn", "Reading Seller template...");
      sellerTemplateBytes = await fileToUint8Array(file);
      setPill(sellerPill, "ok", file.name);
      logLine("Seller template loaded.");
      refreshGenerateButton();
    } catch (err) {
      console.error(err);
      sellerTemplateBytes = null;
      setPill(sellerPill, "bad", "Seller template failed");
      setStatus("Seller template read failed:\n" + (err?.message || String(err)));
      refreshGenerateButton();
    }
  }

  // ------------------------------------------------------------
  // DOCX render
  // ------------------------------------------------------------
  function renderDocxFromTemplate(templateBytes, dataObj) {
    const PizZipRef = getPizZip();
    const DocxRef = getDocxtemplater();

    // clone: docxtemplater mutates zip state
    const bytes = templateBytes.slice(0);

    const zip = new PizZipRef(bytes);
    const doc = new DocxRef(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: CONFIG.DELIMS, // {{Tag}}
    });

    doc.render(dataObj);
    return doc.getZip().generate({ type: "blob" });
  }

  // ------------------------------------------------------------
  // Generate ZIP
  // ------------------------------------------------------------
  async function generateZip() {
    if (!(csvRows.length > 0 && buyerTemplateBytes && sellerTemplateBytes)) {
      setStatus("Upload CSV + Buyer template + Seller template first.");
      return;
    }

    setStatus("Generating DOCX files and building ZIP...");
    genBtn.disabled = true;

    const zip = new JSZip();
    const buyerFolder = zip.folder(CONFIG.ZIP_FOLDERS.buyer);
    const sellerFolder = zip.folder(CONFIG.ZIP_FOLDERS.seller);

    const usedNames = new Set();
    let okRows = 0;
    let failRows = 0;

    for (let i = 0; i < csvRows.length; i++) {
      const row = csvRows[i];

      try {
        // Your templates use both {{Shrink}} and seller uses {{shrink}}.
        const data = {
          ...row,
          Shrink: row["Shrink"],
          shrink: row["Shrink"],
        };

        const contractNo = sanitizeFilePart(row["Contract #"]);
        const buyerName = sanitizeFilePart(row["Buyer"]);
        const consignor = sanitizeFilePart(row["Consignor"]);

        // Naming rules YOU specified:
        // Seller: Consignor Name-Contract Number
        // Buyer : Contract Number-Buyer Name
        let sellerFile = `${consignor}-${contractNo}.docx`;
        let buyerFile = `${contractNo}-${buyerName}.docx`;

        sellerFile = dedupeName(sellerFile, usedNames);
        buyerFile = dedupeName(buyerFile, usedNames);

        const buyerBlob = renderDocxFromTemplate(buyerTemplateBytes, data);
        const sellerBlob = renderDocxFromTemplate(sellerTemplateBytes, data);

        buyerFolder.file(buyerFile, buyerBlob);
        sellerFolder.file(sellerFile, sellerBlob);

        okRows++;
      } catch (err) {
        failRows++;
        logLine(`Row ${i + 1} failed: ${err?.message || String(err)}`);
      }
    }

    const zipName = `CMS Contracts - ${new Date().toISOString().slice(0, 10)}.zip`;

    try {
      const outBlob = await zip.generateAsync({ type: "blob" });
      saveAs(outBlob, zipName);

      setStatus(
        `Done.\n` +
          `Rows processed: ${csvRows.length}\n` +
          `Rows succeeded: ${okRows}\n` +
          `Rows failed: ${failRows}\n` +
          `Contracts generated: ${okRows * 2} (Buyer+Seller per successful row)\n\n` +
          `Downloaded: ${zipName}\n\n` +
          (failRows ? "Scroll status log for row-level errors." : "No errors.")
      );
    } catch (err) {
      console.error(err);
      setStatus("ZIP build failed:\n" + (err?.message || String(err)));
    }

    refreshGenerateButton();
  }

  // ------------------------------------------------------------
  // Wire events
  // ------------------------------------------------------------
  function wireEvents() {
    // PIN
    pinBtn.addEventListener("click", handlePinSubmit);
    pinInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") handlePinSubmit();
    });

    // Exit/Clear = reload (clears memory)
    exitBtn.addEventListener("click", () => window.location.reload());

    // Drag/drop zones
    wireDropZone(dropCsv, async (file) => {
      if (!file.name.toLowerCase().endsWith(".csv")) {
        setPill(csvPill, "bad", "Not a CSV");
        setStatus("That file is not a .csv. Please drop the auction results CSV.");
        return;
      }
      await handleCsvFile(file);
    });

    wireDropZone(dropBuyer, async (file) => {
      if (!file.name.toLowerCase().endsWith(".docx")) {
        setPill(buyerPill, "bad", "Not a DOCX");
        setStatus("Buyer template must be a .docx file.");
        return;
      }
      await handleBuyerTemplate(file);
    });

    wireDropZone(dropSeller, async (file) => {
      if (!file.name.toLowerCase().endsWith(".docx")) {
        setPill(sellerPill, "bad", "Not a DOCX");
        setStatus("Seller template must be a .docx file.");
        return;
      }
      await handleSellerTemplate(file);
    });

    // Choose file buttons
    csvPickBtn.addEventListener("click", () => csvPicker.click());
    buyerPickBtn.addEventListener("click", () => buyerPicker.click());
    sellerPickBtn.addEventListener("click", () => sellerPicker.click());

    // Pickers
    csvPicker.addEventListener("change", async (e) => {
      const f = e.target.files?.[0];
      if (f) await handleCsvFile(f);
      e.target.value = "";
    });

    buyerPicker.addEventListener("change", async (e) => {
      const f = e.target.files?.[0];
      if (f) await handleBuyerTemplate(f);
      e.target.value = "";
    });

    sellerPicker.addEventListener("change", async (e) => {
      const f = e.target.files?.[0];
      if (f) await handleSellerTemplate(f);
      e.target.value = "";
    });

    // Generate
    genBtn.addEventListener("click", generateZip);
  }

  // ------------------------------------------------------------
  // Init
  // ------------------------------------------------------------
  function init() {
    if (!assertDom()) return;
    if (!requireLibs()) return;

    setPill(csvPill, "warn", "No CSV loaded");
    setPill(buyerPill, "warn", "No Buyer template loaded");
    setPill(sellerPill, "warn", "No Seller template loaded");
    refreshGenerateButton();

    setStatus("Locked. Enter PIN to begin.");
    console.log("CMS Contract Generator loaded.");
    wireEvents();
  }

  init();
})();
