/* CMS Contract Generator — FULL app.js (copy/replace)
   - PIN gate: 0623
   - Upload CSV + Buyer DOCX template + Seller DOCX template EACH SESSION (templates not stored on web)
   - For each CSV row: generate Buyer + Seller editable DOCX using docxtemplater with {{ }} delimiters
   - ZIP output:
       /Buyer Contracts/{Contract#}-{Buyer}.docx
       /Seller Contracts/{Consignor}-{Contract#}.docx
   - Matches the IDs in your latest index.html (post-auction style)
*/

(() => {
  "use strict";

  const CONFIG = {
    PIN: "0623",
    ZIP_FOLDERS: {
      buyer: "Buyer Contracts",
      seller: "Seller Contracts",
    },
    // Required CSV headers that MUST exist
    REQUIRED_COLS: [
      "Contract #",
      "Consignor",
      "Buyer",
      "Lot Number #2",
      "Head Count",
    ],
    // Your DOCX templates currently use {{Tag}} style
    DELIMS: { start: "{{", end: "}}" },
  };

  // ------------------------------
  // IMPORTANT: Global drag/drop prevent
  // Fixes the "I dropped CSV but it still says none uploaded" issue in many browsers
  // ------------------------------
  window.addEventListener("dragover", (e) => e.preventDefault(), false);
  window.addEventListener("drop", (e) => e.preventDefault(), false);

  // ------------------------------
  // DOM helpers
  // ------------------------------
  const $ = (id) => document.getElementById(id);

  const pinScreen = $("pinScreen");
  const appScreen = $("appScreen");
  const pinInput  = $("pinInput");
  const pinBtn    = $("pinBtn");
  const pinMsg    = $("pinMsg");

  const exitBtn   = $("exitBtn");

  const dropCsv    = $("dropCsv");
  const dropBuyer  = $("dropBuyer");
  const dropSeller = $("dropSeller");

  const csvPicker   = $("csvPicker");
  const buyerPicker = $("buyerPicker");
  const sellerPicker= $("sellerPicker");

  const csvPickBtn   = $("csvPickBtn");
  const buyerPickBtn = $("buyerPickBtn");
  const sellerPickBtn= $("sellerPickBtn");

  const csvPill    = $("csvPill");
  const buyerPill  = $("buyerPill");
  const sellerPill = $("sellerPill");

  const genBtn   = $("genBtn");
  const statusEl = $("status");

  // ------------------------------
  // Validate required libraries exist
  // ------------------------------
  function requireGlobals() {
    const missing = [];
    if (typeof Papa === "undefined") missing.push("PapaParse (Papa)");
    const PizZipRef = window.PizZip || window.pizzip || window.Pizzip;
    if (!PizZipRef) missing.push("PizZip");
    const DocxRef = window.docxtemplater || window.Docxtemplater || window.DocxTemplater;
    if (!DocxRef) missing.push("docxtemplater");
    if (typeof JSZip === "undefined") missing.push("JSZip");
    if (typeof saveAs === "undefined") missing.push("FileSaver (saveAs)");

    if (missing.length) {
      setStatus(
        "Missing required libraries:\n" +
        missing.map(m => `- ${m}`).join("\n") +
        "\n\nFix: confirm these files exist in /lib and are loaded BEFORE app.js:\n" +
        "- pizzip.min.js\n- docxtemplater.min.js\n- jszip.min.js\n- FileSaver.min.js\n- papa.min.js"
      );
      console.error("Missing libs:", missing);
      return false;
    }
    return true;
  }

  // ------------------------------
  // State
  // ------------------------------
  let csvRows = [];
  let buyerTemplateBytes = null;   // Uint8Array
  let sellerTemplateBytes = null;  // Uint8Array

  // ------------------------------
  // UI helpers
  // ------------------------------
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

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  function addDragUI(el, on) {
    el.classList.toggle("dragover", !!on);
  }

  async function fileToUint8Array(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(new Uint8Array(r.result));
      r.onerror = () => reject(new Error("Failed to read file."));
      r.readAsArrayBuffer(file);
    });
  }

  // ------------------------------
  // PIN gate
  // ------------------------------
  function unlock() {
    pinScreen.classList.add("hide");
    appScreen.classList.remove("hide");
    setStatus("Unlocked. Upload CSV + Buyer template + Seller template, then click Generate Contracts ZIP.");
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

  pinBtn?.addEventListener("click", handlePinSubmit);
  pinInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") handlePinSubmit();
  });

  // Exit/Clear = reload page (clears memory)
  exitBtn?.addEventListener("click", () => window.location.reload());

  // ------------------------------
  // CSV handling
  // ------------------------------
  async function handleCsvFile(file) {
    setStatus(`Reading CSV: ${file.name} ...`);

    const text = await file.text();
    const parsed = Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
    });

    if (parsed.errors && parsed.errors.length) {
      setPill(csvPill, "bad", "CSV parse error");
      setStatus(
        "CSV parse errors:\n" +
        parsed.errors.map(e => `${e.message} (row ${e.row})`).join("\n")
      );
      csvRows = [];
      refreshGenerateButton();
      return;
    }

    const headers = parsed.meta?.fields || [];
    const missing = CONFIG.REQUIRED_COLS.filter(c => !headers.includes(c));

    if (missing.length) {
      setPill(csvPill, "bad", "Missing columns");
      setStatus(
        "CSV is missing required columns:\n" +
        missing.map(m => `- ${m}`).join("\n") +
        "\n\nFix: your CSV headers must match exactly (including spaces/#)."
      );
      csvRows = [];
      refreshGenerateButton();
      return;
    }

    const rows = (parsed.data || []).filter(r =>
      Object.values(r).some(v => String(v ?? "").trim() !== "")
    );

    csvRows = rows;
    setPill(csvPill, "ok", `${file.name} (${rows.length} rows)`);
    setStatus(`CSV loaded: ${rows.length} rows.\nNow upload BOTH DOCX templates.`);
    refreshGenerateButton();
  }

  // ------------------------------
  // Template handling
  // ------------------------------
  async function handleBuyerTemplate(file) {
    setStatus(`Reading Buyer template: ${file.name} ...`);
    buyerTemplateBytes = await fileToUint8Array(file);
    setPill(buyerPill, "ok", file.name);
    logLine("Buyer template loaded.");
    refreshGenerateButton();
  }

  async function handleSellerTemplate(file) {
    setStatus(`Reading Seller template: ${file.name} ...`);
    sellerTemplateBytes = await fileToUint8Array(file);
    setPill(sellerPill, "ok", file.name);
    logLine("Seller template loaded.");
    refreshGenerateButton();
  }

  // ------------------------------
  // Drag/drop wiring
  // ------------------------------
  function wireDropZone(el, acceptFn) {
    if (!el) return;

    ["dragenter", "dragover", "dragleave", "drop"].forEach(evt => {
      el.addEventListener(evt, preventDefaults, false);
    });

    ["dragenter", "dragover"].forEach(evt => {
      el.addEventListener(evt, () => addDragUI(el, true), false);
    });

    ["dragleave", "drop"].forEach(evt => {
      el.addEventListener(evt, () => addDragUI(el, false), false);
    });

    el.addEventListener("drop", async (e) => {
      const file = e.dataTransfer?.files?.[0];
      if (!file) return;
      setStatus(`Dropped file: ${file.name}`);
      await acceptFn(file);
    });
  }

  wireDropZone(dropCsv, async (file) => {
    if (!file.name.toLowerCase().endsWith(".csv")) {
      setPill(csvPill, "bad", "Not a CSV");
      setStatus("That file is not a .csv. Please upload the auction results CSV.");
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
  csvPickBtn?.addEventListener("click", () => csvPicker?.click());
  buyerPickBtn?.addEventListener("click", () => buyerPicker?.click());
  sellerPickBtn?.addEventListener("click", () => sellerPicker?.click());

  csvPicker?.addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (f) await handleCsvFile(f);
    e.target.value = "";
  });

  buyerPicker?.addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (f) await handleBuyerTemplate(f);
    e.target.value = "";
  });

  sellerPicker?.addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (f) await handleSellerTemplate(f);
    e.target.value = "";
  });

  // ------------------------------
  // DOCX render
  // ------------------------------
  function renderDocxFromTemplate(templateBytes, dataObj) {
    const PizZipRef = window.PizZip || window.pizzip || window.Pizzip;
    const DocxRef = window.docxtemplater || window.Docxtemplater || window.DocxTemplater;

    // clone because docxtemplater mutates
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

  // ------------------------------
  // Generate ZIP
  // ------------------------------
  genBtn?.addEventListener("click", async () => {
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
        // Provide BOTH Shrink and shrink because your templates use both variants
        const data = {
          ...row,
          Shrink: row["Shrink"],
          shrink: row["Shrink"],
        };

        const contractNo = sanitizeFilePart(row["Contract #"]);
        const buyerName  = sanitizeFilePart(row["Buyer"]);
        const consignor  = sanitizeFilePart(row["Consignor"]);

        // Naming rules (yours)
        let sellerFile = `${consignor}-${contractNo}.docx`;
        let buyerFile  = `${contractNo}-${buyerName}.docx`;

        sellerFile = dedupeName(sellerFile, usedNames);
        buyerFile  = dedupeName(buyerFile, usedNames);

        const buyerBlob = renderDocxFromTemplate(buyerTemplateBytes, data);
        const sellerBlob = renderDocxFromTemplate(sellerTemplateBytes, data);

        buyerFolder.file(buyerFile, buyerBlob);
        sellerFolder.file(sellerFile, sellerBlob);

        okRows++;
      } catch (err) {
        failRows++;
        logLine(`Row ${i + 1} failed: ${err?.message || err}`);
      }
    }

    const zipName = `CMS Contracts - ${new Date().toISOString().slice(0,10)}.zip`;

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
    } catch (e) {
      setStatus("ZIP build failed: " + (e?.message || e));
    }

    refreshGenerateButton();
  });

  // ------------------------------
  // Init
  // ------------------------------
  setPill(csvPill, "warn", "No CSV loaded");
  setPill(buyerPill, "warn", "No Buyer template loaded");
  setPill(sellerPill, "warn", "No Seller template loaded");
  refreshGenerateButton();

  // Verify libs after load
  if (!requireGlobals()) return;

  console.log("CMS Contract Generator loaded.");
})();
