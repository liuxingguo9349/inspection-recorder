(function () {
  "use strict";

  const DB_NAME = "inspection-recorder-db";
  const DB_VERSION = 1;
  const STORE_NAME = "records";
  const SELECTION_KEY = "inspection-recorder-selection";
  const MOBILE_TAB_KEY = "inspection-recorder-mobile-tab";
  const TEMPLATE_URL = "./assets/template.xlsx";
  const TYPES = ["市政", "市容", "环卫", "绿化", "执法", "其他"];
  const IMAGE_MAX_EDGE = 1800;
  const IMAGE_JPEG_QUALITY = 0.82;
  const IMAGE_COMPRESSION_THRESHOLD = 1.5 * 1024 * 1024;
  const EXPORT_TITLE_PREFIX = "达沃斯重点保障区域市容整治包保检查问题";
  const EXPORT_HEADERS = ["序号", "问题类型", "问题简要描述", "问题照片", "整改照片", "备注"];
  const EXPORT_LAYOUT = {
    columnCount: 6,
    headerRow: 2,
    firstDataRow: 3,
    issuePhotoColumn: 4,
    photoPaddingPx: 2,
  };

  const state = {
    db: null,
    records: [],
    filteredRecords: [],
    selectedIds: new Set(loadSelection()),
    draftIssuePhotoDataUrl: "",
    draftIssuePhotoName: "",
    draftIssuePhotoSize: 0,
    toastTimer: 0,
  };

  const els = {
    form: document.getElementById("issueForm"),
    recordId: document.getElementById("recordId"),
    typeSelect: document.getElementById("typeSelect"),
    recordTime: document.getElementById("recordTime"),
    description: document.getElementById("description"),
    remarks: document.getElementById("remarks"),
    issuePhotoInput: document.getElementById("issuePhotoInput"),
    issuePhotoPreview: document.getElementById("issuePhotoPreview"),
    issuePhotoEmpty: document.getElementById("issuePhotoEmpty"),
    removeIssuePhotoBtn: document.getElementById("removeIssuePhotoBtn"),
    saveBtn: document.getElementById("saveBtn"),
    resetBtn: document.getElementById("resetBtn"),
    newRecordBtn: document.getElementById("newRecordBtn"),
    editorTitle: document.getElementById("editorTitle"),
    recordList: document.getElementById("recordList"),
    recordCount: document.getElementById("recordCount"),
    statusLine: document.getElementById("statusLine"),
    exportBtn: document.getElementById("exportBtn"),
    selectAllBtn: document.getElementById("selectAllBtn"),
    clearSelectBtn: document.getElementById("clearSelectBtn"),
    searchInput: document.getElementById("searchInput"),
    filterType: document.getElementById("filterType"),
    backupBtn: document.getElementById("backupBtn"),
    restoreInput: document.getElementById("restoreInput"),
    mobileExportBtn: document.getElementById("mobileExportBtn"),
    mobileTabButtons: document.querySelectorAll("[data-mobile-tab]"),
    toast: document.getElementById("toast"),
  };

  document.addEventListener("DOMContentLoaded", init);

  async function init() {
    renderIcons();
    setDefaultTime();
    bindEvents();
    state.db = await openDatabase();
    await requestPersistentStorage();
    await refreshRecords();
    registerServiceWorker();
  }

  function bindEvents() {
    els.form.addEventListener("submit", handleSave);
    els.resetBtn.addEventListener("click", clearForm);
    els.newRecordBtn.addEventListener("click", clearForm);
    els.issuePhotoInput.addEventListener("change", (event) => handlePhotoChange(event, "issue"));
    els.removeIssuePhotoBtn.addEventListener("click", () => clearDraftPhoto("issue"));
    els.recordList.addEventListener("click", handleRecordListClick);
    els.recordList.addEventListener("change", handleRecordListChange);
    els.searchInput.addEventListener("input", renderRecords);
    els.filterType.addEventListener("change", renderRecords);
    els.selectAllBtn.addEventListener("click", selectVisibleRecords);
    els.clearSelectBtn.addEventListener("click", clearSelection);
    els.exportBtn.addEventListener("click", exportSelectedRecords);
    els.mobileExportBtn.addEventListener("click", exportSelectedRecords);
    els.mobileTabButtons.forEach((button) => {
      button.addEventListener("click", () => switchMobileTab(button.dataset.mobileTab));
    });
    els.backupBtn.addEventListener("click", exportBackup);
    els.restoreInput.addEventListener("change", importBackup);
  }

  function openDatabase() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          const store = db.createObjectStore(STORE_NAME, { keyPath: "id" });
          store.createIndex("updatedAt", "updatedAt", { unique: false });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  function transaction(mode = "readonly") {
    return state.db.transaction(STORE_NAME, mode).objectStore(STORE_NAME);
  }

  function getAllRecords() {
    return new Promise((resolve, reject) => {
      const request = transaction().getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  }

  function putRecord(record) {
    return new Promise((resolve, reject) => {
      const request = transaction("readwrite").put(record);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  }

  function deleteRecord(id) {
    return new Promise((resolve, reject) => {
      const request = transaction("readwrite").delete(id);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  }

  async function refreshRecords() {
    const records = await getAllRecords();
    state.records = records.map(normalizeRecord).sort((a, b) => dateValue(b.updatedAt) - dateValue(a.updatedAt));
    removeMissingSelections();
    renderRecords();
  }

  async function handleSave(event) {
    event.preventDefault();
    const existing = normalizeRecord(state.records.find((record) => record.id === els.recordId.value) || {});
    const now = new Date().toISOString();
    const id = existing.id || createId();
    const record = {
      id,
      type: TYPES.includes(els.typeSelect.value) ? els.typeSelect.value : "其他",
      description: els.description.value.trim(),
      remarks: els.remarks.value.trim(),
      recordTime: localInputToIso(els.recordTime.value) || now,
      issuePhotoDataUrl: state.draftIssuePhotoDataUrl,
      issuePhotoName: state.draftIssuePhotoName,
      issuePhotoSize: state.draftIssuePhotoSize,
      fixPhotoDataUrl: "",
      fixPhotoName: "",
      fixPhotoSize: 0,
      createdAt: existing.createdAt || now,
      updatedAt: now,
    };

    await putRecord(record);
    state.selectedIds.add(record.id);
    saveSelection();
    clearForm();
    await refreshRecords();
    switchMobileTab("list");
    showToast("记录和照片已保存到浏览器本地库");
  }

  async function handlePhotoChange(event, kind) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      showToast("请选择图片文件");
      event.target.value = "";
      return;
    }

    try {
      showToast("正在处理照片");
      const photo = await preparePhotoFile(file, kind);
      setDraftPhoto(kind, photo.dataUrl, photo.name, photo.size);
      renderPhotoPreview(kind, photo.dataUrl);
      const suffix = photo.compressed ? `，已压缩 ${formatBytes(photo.originalSize)} -> ${formatBytes(photo.size)}` : "";
      showToast(`问题照片已保存到当前记录${suffix}`);
    } catch (error) {
      console.error(error);
      showToast("照片读取失败");
    }
  }

  function setDraftPhoto(kind, dataUrl, name, size) {
    if (kind === "issue") {
      state.draftIssuePhotoDataUrl = dataUrl;
      state.draftIssuePhotoName = name;
      state.draftIssuePhotoSize = size;
    }
  }

  function clearDraftPhoto(kind) {
    setDraftPhoto(kind, "", "", 0);
    els.issuePhotoInput.value = "";
    renderPhotoPreview(kind, "");
  }

  function handleRecordListChange(event) {
    const checkbox = event.target.closest("[data-select-id]");
    if (!checkbox) return;
    if (checkbox.checked) {
      state.selectedIds.add(checkbox.dataset.selectId);
    } else {
      state.selectedIds.delete(checkbox.dataset.selectId);
    }
    saveSelection();
    renderRecords();
  }

  async function handleRecordListClick(event) {
    const button = event.target.closest("button[data-action]");
    if (!button) return;
    const id = button.dataset.id;
    if (!id) return;

    if (button.dataset.action === "edit") {
      editRecord(id);
      return;
    }

    if (button.dataset.action === "download-issue") {
      downloadRecordPhoto(id, "issue");
      return;
    }

    if (button.dataset.action === "delete") {
      const record = state.records.find((item) => item.id === id);
      const ok = window.confirm(`删除「${record?.type || "记录"}」这条记录？`);
      if (!ok) return;
      await deleteRecord(id);
      state.selectedIds.delete(id);
      saveSelection();
      await refreshRecords();
      showToast("记录已删除");
    }
  }

  function editRecord(id) {
    const record = normalizeRecord(state.records.find((item) => item.id === id) || {});
    if (!record.id) return;
    els.recordId.value = record.id;
    els.typeSelect.value = record.type || "其他";
    els.description.value = record.description || "";
    els.remarks.value = record.remarks || "";
    els.recordTime.value = isoToLocalInput(record.recordTime || record.createdAt);
    state.draftIssuePhotoDataUrl = record.issuePhotoDataUrl || "";
    state.draftIssuePhotoName = record.issuePhotoName || "";
    state.draftIssuePhotoSize = record.issuePhotoSize || 0;
    renderPhotoPreview("issue", state.draftIssuePhotoDataUrl);
    els.editorTitle.textContent = "编辑记录";
    els.saveBtn.querySelector("span").textContent = "保存修改";
    switchMobileTab("form");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function clearForm() {
    els.form.reset();
    els.recordId.value = "";
    els.typeSelect.value = "市政";
    state.draftIssuePhotoDataUrl = "";
    state.draftIssuePhotoName = "";
    state.draftIssuePhotoSize = 0;
    els.issuePhotoInput.value = "";
    renderPhotoPreview("issue", "");
    setDefaultTime();
    els.editorTitle.textContent = "新增记录";
    els.saveBtn.querySelector("span").textContent = "保存记录";
    switchMobileTab("form");
  }

  function setDefaultTime() {
    els.recordTime.value = isoToLocalInput(new Date().toISOString());
  }

  function renderPhotoPreview(kind, dataUrl) {
    const preview = els.issuePhotoPreview;
    const empty = els.issuePhotoEmpty;
    if (dataUrl) {
      preview.src = dataUrl;
      preview.hidden = false;
      empty.hidden = true;
      return;
    }
    preview.removeAttribute("src");
    preview.hidden = true;
    empty.hidden = false;
  }

  function renderRecords() {
    const query = els.searchInput.value.trim().toLowerCase();
    const filterType = els.filterType.value;
    const filtered = state.records.filter((record) => {
      const typeOk = filterType === "全部" || record.type === filterType;
      const text = `${record.type || ""} ${record.description || ""} ${record.remarks || ""}`.toLowerCase();
      return typeOk && (!query || text.includes(query));
    });
    state.filteredRecords = filtered;
    els.recordList.replaceChildren();

    if (!filtered.length) {
      const empty = document.createElement("div");
      empty.className = "empty-state";
      empty.textContent = state.records.length ? "没有匹配记录" : "暂无记录";
      els.recordList.append(empty);
    } else {
      filtered.forEach((record) => els.recordList.append(createRecordCard(record)));
    }

    const selectedCount = state.records.filter((record) => state.selectedIds.has(record.id)).length;
    els.recordCount.textContent = `${state.records.length} 条，已选 ${selectedCount} 条`;
    els.statusLine.textContent = `浏览器本地库 ${state.records.length} 条`;
    els.exportBtn.disabled = selectedCount === 0;
    els.mobileExportBtn.disabled = selectedCount === 0;
    updateMobileTabs();
    renderIcons();
  }

  function switchMobileTab(tab) {
    const nextTab = tab === "list" ? "list" : "form";
    document.body.dataset.mobileTab = nextTab;
    localStorage.setItem(MOBILE_TAB_KEY, nextTab);
    updateMobileTabs();
    if (window.matchMedia("(max-width: 720px)").matches) {
      window.scrollTo({ top: 0, behavior: "auto" });
    }
  }

  function updateMobileTabs() {
    const current = document.body.dataset.mobileTab || localStorage.getItem(MOBILE_TAB_KEY) || "form";
    document.body.dataset.mobileTab = current === "list" ? "list" : "form";
    els.mobileTabButtons.forEach((button) => {
      button.classList.toggle("is-active", button.dataset.mobileTab === document.body.dataset.mobileTab);
    });
  }

  function createRecordCard(rawRecord) {
    const record = normalizeRecord(rawRecord);
    const card = document.createElement("article");
    card.className = `record-card${state.selectedIds.has(record.id) ? " is-selected" : ""}`;

    const checkLabel = document.createElement("label");
    checkLabel.className = "record-check";
    const check = document.createElement("input");
    check.type = "checkbox";
    check.checked = state.selectedIds.has(record.id);
    check.dataset.selectId = record.id;
    check.setAttribute("aria-label", "选择记录");
    checkLabel.append(check);

    const thumb = document.createElement("div");
    thumb.className = "record-thumb";
    const thumbDataUrl = record.issuePhotoDataUrl;
    if (thumbDataUrl) {
      const img = document.createElement("img");
      img.src = thumbDataUrl;
      img.alt = "记录照片";
      thumb.append(img);
    } else {
      const empty = document.createElement("div");
      empty.className = "record-thumb-empty";
      const icon = document.createElement("i");
      icon.dataset.lucide = "image";
      empty.append(icon);
      thumb.append(empty);
    }

    const body = document.createElement("div");
    body.className = "record-body";
    const meta = document.createElement("div");
    meta.className = "record-meta";
    const pill = document.createElement("span");
    pill.className = "type-pill";
    pill.textContent = record.type || "其他";
    const photoBadge = document.createElement("span");
    photoBadge.className = "photo-badge";
    photoBadge.textContent = record.issuePhotoDataUrl ? "有问题照片" : "无问题照片";
    const time = document.createElement("span");
    time.className = "time-text";
    time.textContent = formatDisplayTime(record.recordTime || record.createdAt);
    meta.append(pill, photoBadge, time);
    const description = document.createElement("p");
    description.className = "description-text";
    description.textContent = record.description || "未填写描述";
    body.append(meta, description);

    const actions = document.createElement("div");
    actions.className = "record-actions";
    actions.append(
      iconButton("edit-3", "编辑", "edit", record.id),
      iconButton("download", "问题图", "download-issue", record.id, !record.issuePhotoDataUrl),
      iconButton("trash-2", "删除", "delete", record.id, false, "is-danger"),
    );

    card.append(checkLabel, thumb, body, actions);
    return card;
  }

  function iconButton(iconName, label, action, id, disabled = false, extraClass = "") {
    const button = document.createElement("button");
    button.className = `ghost-button ${extraClass}`.trim();
    button.type = "button";
    button.dataset.action = action;
    button.dataset.id = id;
    button.title = label;
    button.disabled = disabled;
    const icon = document.createElement("i");
    icon.dataset.lucide = iconName;
    const text = document.createElement("span");
    text.className = "sr-only";
    text.textContent = label;
    button.append(icon, text);
    return button;
  }

  function selectVisibleRecords() {
    state.filteredRecords.forEach((record) => state.selectedIds.add(record.id));
    saveSelection();
    renderRecords();
  }

  function clearSelection() {
    state.selectedIds.clear();
    saveSelection();
    renderRecords();
  }

  async function exportSelectedRecords() {
    const selected = state.records.filter((record) => state.selectedIds.has(record.id)).map(normalizeRecord);
    if (!selected.length) {
      showToast("先选择要导出的记录");
      return;
    }
    if (!window.ExcelJS) {
      showToast("Excel 导出库未加载");
      return;
    }

    els.exportBtn.disabled = true;
    els.exportBtn.querySelector("span").textContent = "导出中";

    let stage = "准备导出";
    try {
      stage = "读取照片尺寸";
      const dimensionsById = await resolveImageDimensions(selected);
      stage = "加载 Excel 模板";
      let workbook = await loadTemplateWorkbook().catch((error) => {
        console.error(error);
        showToast("模板加载失败，正在使用备用样式");
        return createFallbackWorkbook();
      });
      workbook.creator = "问题记录";
      workbook.created = new Date();
      workbook.modified = new Date();
      let sheet = workbook.worksheets[0];

      const warnings = [];
      stage = "填充 Excel";
      buildWorkbookSheet(workbook, sheet, selected, dimensionsById, warnings);
      stage = "生成 Excel 文件";
      let buffer;
      try {
        buffer = await workbook.xlsx.writeBuffer();
      } catch (writeError) {
        console.error(writeError);
        warnings.push("模板写出失败，已改用备用样式");
        workbook = createFallbackWorkbook();
        sheet = workbook.worksheets[0];
        buildWorkbookSheet(workbook, sheet, selected, dimensionsById, warnings);
        buffer = await workbook.xlsx.writeBuffer();
      }
      const filename = `达沃斯重点保障区域市容整治包保检查问题-${formatFileDate(new Date())}.xlsx`;
      stage = "下载 Excel";
      downloadBlob(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        filename,
      );
      const suffix = warnings.length ? `，${warnings[0]}` : "";
      showToast(`已导出 ${selected.length} 条记录${suffix}`);
    } catch (error) {
      console.error(error);
      showToast(`导出失败：${stage}，${getErrorMessage(error)}`);
    } finally {
      els.exportBtn.querySelector("span").textContent = "导出选中";
      renderRecords();
    }
  }

  async function loadTemplateWorkbook() {
    const response = await fetch(TEMPLATE_URL, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Template load failed: ${response.status}`);
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await response.arrayBuffer());
    return workbook;
  }

  function createFallbackWorkbook() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1", {
      views: [{ showGridLines: false, state: "frozen", ySplit: 2 }],
      pageSetup: {
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        paperSize: 9,
        orientation: "portrait",
        horizontalCentered: true,
        printTitlesRow: "1:2",
      },
    });
    sheet.columns = [
      { width: 6.50909090909091 },
      { width: 12.7818181818182 },
      { width: 14.9454545454545 },
      { width: 16.5090909090909 },
      { width: 16.5090909090909 },
      { width: 16.0363636363636 },
    ];
    sheet.mergeCells("A1:F1");
    sheet.getRow(1).height = 72;
    sheet.getRow(2).height = 33;
    sheet.getRow(3).height = 117;
    applyFixedSheetTypography(sheet, 1);
    return workbook;
  }

  function buildWorkbookSheet(workbook, sheet, selected, dimensionsById, warnings = []) {
    const templateRows = captureTemplateRows(sheet);
    clearTemplateImages(workbook, sheet);
    trimSurplusTemplateRows(sheet, selected.length, templateRows.length);
    applyFixedSheetTypography(sheet, selected.length);

    selected.forEach((record, index) => {
      const rowNumber = EXPORT_LAYOUT.firstDataRow + index;
      const row = sheet.getRow(rowNumber);
      applyTemplateRow(row, templateRows[Math.min(index, templateRows.length - 1)]);
      applyDataRowTypography(row);
      row.getCell(1).value = index + 1;
      row.getCell(2).value = record.type || "其他";
      row.getCell(3).value = record.description || "";
      row.getCell(4).value = "";
      row.getCell(5).value = "";
      row.getCell(6).value = record.remarks || "";

      if (record.issuePhotoDataUrl) {
        try {
          addPhotoToSheet(
            workbook,
            sheet,
            record.issuePhotoDataUrl,
            rowNumber,
            EXPORT_LAYOUT.issuePhotoColumn,
            dimensionsById.get(`${record.id}:issue`),
          );
        } catch (error) {
          console.error(error);
          warnings.push(`第 ${index + 1} 张照片插入失败，已跳过`);
        }
      }
    });

    const lastRow = Math.max(EXPORT_LAYOUT.firstDataRow, EXPORT_LAYOUT.firstDataRow + selected.length - 1);
    sheet.autoFilter = `A${EXPORT_LAYOUT.headerRow}:F${lastRow}`;
    sheet.pageSetup.printArea = `A1:F${lastRow}`;
  }

  function applyFixedSheetTypography(sheet, selectedCount) {
    if (!sheet.getCell("A1").isMerged) {
      try {
        sheet.mergeCells("A1:F1");
      } catch {
        // The template may already carry merge metadata.
      }
    }
    sheet.getRow(1).height = 72;
    sheet.getRow(2).height = 33;
    sheet.getCell("A1").value = buildExportTitle();
    for (let colNumber = 1; colNumber <= EXPORT_LAYOUT.columnCount; colNumber += 1) {
      const titleCell = sheet.getRow(1).getCell(colNumber);
      titleCell.font = { name: "宋体", size: 24, bold: true, color: { argb: "FF000000" }, charset: 134 };
      titleCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      titleCell.border = blackThinBorder();

      const headerCell = sheet.getRow(2).getCell(colNumber);
      headerCell.value = EXPORT_HEADERS[colNumber - 1];
      headerCell.font = { name: "黑体", size: 11, bold: false, color: { argb: "FF000000" }, charset: 134 };
      headerCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      headerCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      headerCell.border = blackThinBorder();
    }

    const lastDataRow = EXPORT_LAYOUT.firstDataRow + Math.max(1, selectedCount) - 1;
    for (let rowNumber = EXPORT_LAYOUT.firstDataRow; rowNumber <= lastDataRow; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      row.height = 117;
      row.hidden = false;
      applyDataRowTypography(row);
    }
  }

  function applyDataRowTypography(row) {
    row.height = 117;
    row.hidden = false;
    for (let colNumber = 1; colNumber <= EXPORT_LAYOUT.columnCount; colNumber += 1) {
      const cell = row.getCell(colNumber);
      cell.font = { name: "宋体", size: 11, bold: false, color: { argb: "FF000000" }, charset: 134 };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cell.border = blackThinBorder();
    }
  }

  function blackThinBorder() {
    return {
      top: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
  }

  function buildExportTitle(date = new Date()) {
    return `${EXPORT_TITLE_PREFIX}\n（${formatMonthDay(date)}）`;
  }

  function addPhotoToSheet(workbook, sheet, dataUrl, rowNumber, excelColumn, dimensions) {
    if (typeof dataUrl !== "string" || !dataUrl.startsWith("data:image/")) {
      throw new Error("照片数据格式不正确");
    }
    const safeDimensions = dimensions || { width: 4, height: 3 };
    const extension = dataUrl.includes("image/png") ? "png" : "jpeg";
    const imageId = workbook.addImage({
      base64: dataUrl,
      extension,
    });
    const boxWidth = columnWidthToPixels(sheet.getColumn(excelColumn).width);
    const boxHeight = rowHeightToPixels(sheet.getRow(rowNumber).height);
    const padding = EXPORT_LAYOUT.photoPaddingPx;
    const maxWidth = boxWidth - padding * 2;
    const maxHeight = boxHeight - padding * 2;
    const fit = fitInside(safeDimensions.width, safeDimensions.height, maxWidth, maxHeight);
    const offsetX = padding + (maxWidth - fit.width) / 2;
    const offsetY = padding + (maxHeight - fit.height) / 2;
    sheet.addImage(imageId, {
      tl: {
        col: excelColumn - 1 + offsetX / boxWidth,
        row: rowNumber - 1 + offsetY / boxHeight,
      },
      ext: {
        width: fit.width,
        height: fit.height,
      },
      editAs: "oneCell",
    });
  }

  function captureTemplateRows(sheet) {
    const rows = [];
    const lastRow = Math.max(sheet.rowCount, EXPORT_LAYOUT.firstDataRow);
    for (let rowNumber = EXPORT_LAYOUT.firstDataRow; rowNumber <= lastRow; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      const cells = [];
      for (let colNumber = 1; colNumber <= EXPORT_LAYOUT.columnCount; colNumber += 1) {
        const cell = row.getCell(colNumber);
        cells.push({
          style: clonePlain(cell.style),
          numFmt: cell.numFmt,
          alignment: clonePlain(cell.alignment),
          font: clonePlain(cell.font),
          fill: clonePlain(cell.fill),
          border: clonePlain(cell.border),
        });
      }
      rows.push({
        height: row.height,
        cells,
      });
    }
    return rows.length ? rows : [{ height: 116, cells: [] }];
  }

  function applyTemplateRow(row, template) {
    row.height = template.height;
    row.hidden = false;
    for (let colNumber = 1; colNumber <= EXPORT_LAYOUT.columnCount; colNumber += 1) {
      const cell = row.getCell(colNumber);
      const source = template.cells[colNumber - 1];
      if (source) {
        cell.style = clonePlain(source.style);
        if (source.numFmt) cell.numFmt = source.numFmt;
        if (source.alignment) cell.alignment = clonePlain(source.alignment);
        if (source.font) cell.font = clonePlain(source.font);
        if (source.fill) cell.fill = clonePlain(source.fill);
        if (source.border) cell.border = clonePlain(source.border);
      }
      cell.value = null;
    }
  }

  function trimSurplusTemplateRows(sheet, selectedCount, templateRowCount) {
    const firstHiddenRow = EXPORT_LAYOUT.firstDataRow + selectedCount;
    const lastTemplateRow = EXPORT_LAYOUT.firstDataRow + templateRowCount - 1;
    for (let rowNumber = firstHiddenRow; rowNumber <= lastTemplateRow; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      row.hidden = true;
      row.height = 0.1;
    }
  }

  function clearTemplateImages(workbook, sheet) {
    if (Array.isArray(sheet._media)) sheet._media = [];
    if (Array.isArray(workbook.media)) workbook.media = [];
  }

  function columnWidthToPixels(width) {
    return Math.max(1, Math.floor((width || 8.43) * 7 + 5));
  }

  function rowHeightToPixels(height) {
    return Math.max(1, Math.round((height || 15) * 4 / 3));
  }

  async function resolveImageDimensions(records) {
    const entries = [];
    await Promise.all(
      records.map(async (record) => {
        if (record.issuePhotoDataUrl) {
          entries.push([`${record.id}:issue`, await loadImageDimensions(record.issuePhotoDataUrl)]);
        }
      }),
    );
    return new Map(entries);
  }

  function loadImageDimensions(dataUrl) {
    return new Promise((resolve) => {
      const image = new Image();
      image.onload = () => {
        resolve({
          width: image.naturalWidth || 4,
          height: image.naturalHeight || 3,
        });
      };
      image.onerror = () => resolve({ width: 4, height: 3 });
      image.src = dataUrl;
    });
  }

  function fitInside(width, height, maxWidth, maxHeight) {
    const safeWidth = Math.max(1, width || 1);
    const safeHeight = Math.max(1, height || 1);
    const scale = Math.min(maxWidth / safeWidth, maxHeight / safeHeight);
    return {
      width: Math.max(1, Math.round(safeWidth * scale)),
      height: Math.max(1, Math.round(safeHeight * scale)),
    };
  }

  function clonePlain(value) {
    return value ? JSON.parse(JSON.stringify(value)) : value;
  }

  function getErrorMessage(error) {
    const message = error?.message || String(error || "");
    return message.length > 36 ? `${message.slice(0, 36)}...` : message || "未知错误";
  }

  function exportBackup() {
    const payload = {
      version: 2,
      exportedAt: new Date().toISOString(),
      records: state.records,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: "application/json;charset=utf-8",
    });
    downloadBlob(blob, `问题记录备份-${formatFileDate(new Date())}.json`);
    showToast("备份已生成");
  }

  async function importBackup(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const payload = JSON.parse(text);
      const records = Array.isArray(payload.records) ? payload.records.map(normalizeRecord) : [];
      if (!records.length) {
        showToast("备份文件没有记录");
        return;
      }
      for (const record of records) {
        if (!record.id) record.id = createId();
        await putRecord(record);
      }
      await refreshRecords();
      showToast(`已恢复 ${records.length} 条记录`);
    } catch (error) {
      console.error(error);
      showToast("恢复失败");
    } finally {
      event.target.value = "";
    }
  }

  function downloadRecordPhoto(id, kind) {
    const record = normalizeRecord(state.records.find((item) => item.id === id) || {});
    const dataUrl = record.issuePhotoDataUrl;
    if (!dataUrl) return;
    const extension = dataUrl.includes("image/png") ? "png" : "jpg";
    const prefix = "问题照片";
    const fallback = `${prefix}-${formatFileDate(new Date(record.recordTime || Date.now()))}.${extension}`;
    const name = record.issuePhotoName || fallback;
    downloadDataUrl(dataUrl, name);
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1200);
  }

  function downloadDataUrl(dataUrl, filename) {
    const anchor = document.createElement("a");
    anchor.href = dataUrl;
    anchor.download = filename;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
  }

  async function preparePhotoFile(file, kind) {
    const originalDataUrl = await fileToDataUrl(file);
    const dimensions = await loadImageDimensions(originalDataUrl);
    const shouldCompress =
      file.size > IMAGE_COMPRESSION_THRESHOLD ||
      Math.max(dimensions.width, dimensions.height) > IMAGE_MAX_EDGE ||
      file.type !== "image/jpeg";

    if (!shouldCompress) {
      return {
        dataUrl: originalDataUrl,
        name: file.name || `${kind}-${Date.now()}.jpg`,
        size: file.size || dataUrlByteLength(originalDataUrl),
        originalSize: file.size || dataUrlByteLength(originalDataUrl),
        compressed: false,
      };
    }

    const compressed = await compressImageDataUrl(originalDataUrl, dimensions);
    const compressedSize = dataUrlByteLength(compressed);
    const originalSize = file.size || dataUrlByteLength(originalDataUrl);
    if (compressedSize >= originalSize && Math.max(dimensions.width, dimensions.height) <= IMAGE_MAX_EDGE) {
      return {
        dataUrl: originalDataUrl,
        name: file.name || `${kind}-${Date.now()}.jpg`,
        size: originalSize,
        originalSize,
        compressed: false,
      };
    }

    return {
      dataUrl: compressed,
      name: replaceImageExtension(file.name || `${kind}-${Date.now()}.jpg`, "jpg"),
      size: compressedSize,
      originalSize,
      compressed: true,
    };
  }

  async function compressImageDataUrl(dataUrl, dimensions) {
    const image = await loadImageElement(dataUrl);
    const scale = Math.min(1, IMAGE_MAX_EDGE / Math.max(dimensions.width, dimensions.height));
    const targetWidth = Math.max(1, Math.round(dimensions.width * scale));
    const targetHeight = Math.max(1, Math.round(dimensions.height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = targetWidth;
    canvas.height = targetHeight;
    const ctx = canvas.getContext("2d", { alpha: false });
    if (!ctx) throw new Error("Canvas is not available");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, targetWidth, targetHeight);
    ctx.drawImage(image, 0, 0, targetWidth, targetHeight);
    const blob = await new Promise((resolve, reject) => {
      canvas.toBlob(
        (result) => {
          if (result) resolve(result);
          else reject(new Error("Image compression failed"));
        },
        "image/jpeg",
        IMAGE_JPEG_QUALITY,
      );
    });
    return fileToDataUrl(blob);
  }

  function loadImageElement(dataUrl) {
    return new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error("Image load failed"));
      image.src = dataUrl;
    });
  }

  function fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });
  }

  function dataUrlByteLength(dataUrl) {
    const base64 = String(dataUrl).split(",")[1] || "";
    return Math.floor((base64.length * 3) / 4);
  }

  function replaceImageExtension(name, extension) {
    const safeName = name || `photo-${Date.now()}.${extension}`;
    return safeName.replace(/\.[^.\\/]+$/, "") + `.${extension}`;
  }

  function formatBytes(bytes) {
    const safeBytes = Math.max(0, bytes || 0);
    if (safeBytes >= 1024 * 1024) return `${(safeBytes / 1024 / 1024).toFixed(1)}MB`;
    if (safeBytes >= 1024) return `${Math.round(safeBytes / 1024)}KB`;
    return `${safeBytes}B`;
  }

  function normalizeRecord(record) {
    const issuePhotoDataUrl = record.issuePhotoDataUrl || record.photoDataUrl || "";
    return {
      id: record.id || "",
      type: TYPES.includes(record.type) ? record.type : "其他",
      description: record.description || "",
      remarks: record.remarks || "",
      recordTime: record.recordTime || record.createdAt || new Date().toISOString(),
      issuePhotoDataUrl,
      issuePhotoName: record.issuePhotoName || record.photoName || "",
      issuePhotoSize: record.issuePhotoSize || record.photoSize || 0,
      fixPhotoDataUrl: "",
      fixPhotoName: "",
      fixPhotoSize: 0,
      createdAt: record.createdAt || new Date().toISOString(),
      updatedAt: record.updatedAt || record.createdAt || new Date().toISOString(),
    };
  }

  function formatDisplayTime(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    const hh = String(date.getHours()).padStart(2, "0");
    const mm = String(date.getMinutes()).padStart(2, "0");
    return `${y}-${m}-${d} ${hh}:${mm}`;
  }

  function formatFileDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    const hh = String(date.getHours()).padStart(2, "0");
    const mm = String(date.getMinutes()).padStart(2, "0");
    return `${y}${m}${d}-${hh}${mm}`;
  }

  function formatMonthDay(date) {
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${m}${d}`;
  }

  function isoToLocalInput(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    const hh = String(date.getHours()).padStart(2, "0");
    const mm = String(date.getMinutes()).padStart(2, "0");
    return `${y}-${m}-${d}T${hh}:${mm}`;
  }

  function localInputToIso(value) {
    if (!value) return "";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    return date.toISOString();
  }

  function dateValue(value) {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? 0 : date.getTime();
  }

  function createId() {
    if (window.crypto?.randomUUID) return crypto.randomUUID();
    return `record-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function loadSelection() {
    try {
      const raw = localStorage.getItem(SELECTION_KEY);
      const parsed = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  function saveSelection() {
    localStorage.setItem(SELECTION_KEY, JSON.stringify([...state.selectedIds]));
  }

  function removeMissingSelections() {
    const ids = new Set(state.records.map((record) => record.id));
    for (const id of [...state.selectedIds]) {
      if (!ids.has(id)) state.selectedIds.delete(id);
    }
    saveSelection();
  }

  async function requestPersistentStorage() {
    if (!navigator.storage?.persist) return;
    try {
      await navigator.storage.persist();
    } catch {
      // Some browsers expose the API but disallow prompts in file contexts.
    }
  }

  function registerServiceWorker() {
    if (!("serviceWorker" in navigator)) return;
    if (!window.isSecureContext && location.hostname !== "localhost") return;
    navigator.serviceWorker.register("./service-worker.js").catch(() => {});
  }

  function showToast(message) {
    window.clearTimeout(state.toastTimer);
    els.toast.textContent = message;
    els.toast.classList.add("is-visible");
    state.toastTimer = window.setTimeout(() => {
      els.toast.classList.remove("is-visible");
    }, 2400);
  }

  function renderIcons() {
    if (window.lucide?.createIcons) {
      window.lucide.createIcons({
        attrs: {
          "stroke-width": 1.8,
        },
      });
    }
  }
})();
