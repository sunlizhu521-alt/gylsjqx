// Files live only in this page session; never read or write persisted uploads.
const sessionFiles = new Map();

const slots = [
  { id: "file-1", label: "Dim-YL医疗器械商品分类" },
  { id: "file-2", label: "Dim-采购部分工明细" },
  { id: "file-3", label: "Dim-花名册（姓名&一级部门）" },
  { id: "file-4", label: "备货需求表" },
];

const TABLE_FILE_ACCEPT = ".xlsx,.xls,.xlsm,.csv";
const DETAIL_PAGE_SIZE = 20;
const SLOT_IDS = {
  category: "file-1",
  purchaseDivision: "file-2",
  roster: "file-3",
  demand: "file-4",
};
const FILTER_DEFINITIONS = [
  { key: "businessUnit", label: "事业部" },
  { key: "productLine", label: "销售产品线" },
  { key: "salesSeries", label: "销售系列" },
  { key: "purchaseGroup", label: "采购组" },
  { key: "buyer", label: "采购订单下单人" },
  { key: "supplierShortName", label: "供应商简称" },
];
const LINKED_FILTER_GROUPS = [
  ["productLine", "salesSeries"],
  ["purchaseGroup", "buyer"],
];
const DETAIL_COLUMNS = [
  { key: "buyer", label: "采购单订单下单人" },
  { key: "businessUnit", label: "事业部" },
  { key: "applicant", label: "申请人" },
  { key: "supplierShortName", label: "供应商简称" },
  { key: "materialCode", label: "物料编码" },
  { key: "sku", label: "SKU" },
  { key: "materialName", label: "物料名称" },
  { key: "quantity", label: "数量" },
  { key: "oaProcessNo", label: "OA备货流程号" },
  { key: "purchaseEntity", label: "采购主体" },
  { key: "materialCodeValid", label: "采购分工明细是否存在" },
  { key: "supplierDeliveryDate", label: "供应商反馈交期" },
  { key: "purchaseOrderNo", label: "采购单订单号" },
  { key: "requiredReadyDate", label: "要求货好时间" },
  { key: "minimumOrderQuantity", label: "起订量" },
  { key: "minimumOrderStatus", label: "起订量是否满足" },
  { key: "stockReason", label: "备货原因" },
];
const TABLE_FIELD_ALIASES = {
  productLine: ["销售产品线", "产品线", "一级产品线", "销售线"],
  salesSeries: ["销售系列", "系列", "产品系列", "销售产品系列"],
  purchaseGroup: ["采购分组", "采购组", "采购组别", "采购分组名称"],
  demandApplicant: ["创建人", "申请人", "OA申请人", "流程申请人"],
  demandMaterialCode: ["识别码", "物料编码", "物料编号", "商品编码", "存货编码", "产品编码"],
  demandProcessNo: ["流程号", "流程编号", "OA备货流程号", "OA流程号", "备货流程号", "备货需求流程号", "OA编号", "OA单号", "单号"],
  demandStockReason: ["备货原因", "申请原因", "需求原因", "原因"],
  demandPurchaseEntity: ["采购主体", "采购公司", "采购组织", "采购方", "主体"],
  demandQuantity: ["数量"],
  demandRequiredReadyDate: ["要求货好时间", "要求货好日期", "货好时间", "货好日期", "要求到货时间", "要求到货日期"],
  divisionPurchaseGroup: ["采购组", "采购分组", "采购组别", "采购分组名称"],
  divisionMaterialCode: ["物料编码", "识别码", "物料编号", "商品编码", "存货编码", "产品编码"],
  divisionBuyer: ["采购单订单下单人", "采购订单下单人", "采购下单人", "采购单下单人", "订单下单人", "下单人", "采购员", "采购负责人", "采购对接人", "采购组对接人"],
  divisionSupplier: ["供应商简称", "供应商", "供应商名称", "供应商短名称"],
};

const els = {
  slotGrid: document.querySelector("#slotGrid"),
  applyAllButton: document.querySelector("#applyAllButton"),
  clearCacheButton: document.querySelector("#clearCacheButton"),
  libraryState: document.querySelector("#libraryState"),
  sourceNote: document.querySelector("#librarySourceNote"),
  slotCount: document.querySelector("#slotCount"),
  uploadedCount: document.querySelector("#uploadedCount"),
  appliedCount: document.querySelector("#appliedCount"),
  beihuoTotal: document.querySelector("#beihuoTotal"),
  quantityTotal: document.querySelector("#quantityTotal"),
  supplierCount: document.querySelector("#supplierCount"),
  filterToolbar: document.querySelector("#filterToolbar"),
  clearFiltersButton: document.querySelector("#clearFiltersButton"),
  downloadDetailButton: document.querySelector("#downloadDetailButton"),
  detailState: document.querySelector("#detailState"),
  detailTableBody: document.querySelector("#detailTableBody"),
  detailEmpty: document.querySelector("#detailEmpty"),
  paginationControls: document.querySelector("#paginationControls"),
  paginationInfo: document.querySelector("#paginationInfo"),
  previousPageButton: document.querySelector("#previousPageButton"),
  nextPageButton: document.querySelector("#nextPageButton"),
};

const state = {
  records: new Map(),
  demandRows: [],
  filteredRows: [],
  currentPage: 1,
  filterOptions: Object.fromEntries(FILTER_DEFINITIONS.map((filter) => [filter.key, []])),
  filters: new Map(FILTER_DEFINITIONS.map((filter) => [filter.key, new Set()])),
};

async function init() {
  bindEvents();
  renderFilterControls();
  await refresh();
}

function bindEvents() {
  els.slotGrid?.addEventListener("change", async (event) => {
    const input = event.target.closest("[data-upload]");
    if (!input) return;
    await savePendingFile(input.dataset.upload, input.files?.[0]);
    input.value = "";
  });

  els.slotGrid?.addEventListener("click", async (event) => {
    const applyButton = event.target.closest("[data-apply]");
    if (applyButton) {
      await applySlot(applyButton.dataset.apply);
      return;
    }

    const deleteButton = event.target.closest("[data-delete]");
    if (deleteButton) {
      await deleteSlot(deleteButton.dataset.delete);
    }
  });

  els.slotGrid?.addEventListener("dragover", (event) => {
    const card = event.target.closest("[data-drop]");
    if (!card) return;
    event.preventDefault();
    card.classList.add("drag-over");
  });

  els.slotGrid?.addEventListener("dragleave", (event) => {
    const card = event.target.closest("[data-drop]");
    if (!card || card.contains(event.relatedTarget)) return;
    card.classList.remove("drag-over");
  });

  els.slotGrid?.addEventListener("drop", async (event) => {
    const card = event.target.closest("[data-drop]");
    if (!card) return;
    event.preventDefault();
    card.classList.remove("drag-over");
    await savePendingFile(card.dataset.drop, event.dataTransfer?.files?.[0]);
  });

  els.applyAllButton?.addEventListener("click", applyAllSlots);
  els.clearCacheButton?.addEventListener("click", clearLibraryCache);
  els.filterToolbar?.addEventListener("click", handleFilterToolbarClick);
  els.filterToolbar?.addEventListener("change", handleFilterToolbarChange);
  els.clearFiltersButton?.addEventListener("click", clearAllFilters);
  els.downloadDetailButton?.addEventListener("click", downloadDetailWorkbook);
  els.previousPageButton?.addEventListener("click", () => changeDetailPage(-1));
  els.nextPageButton?.addEventListener("click", () => changeDetailPage(1));
  document.addEventListener("click", (event) => {
    if (!event.target.closest(".multi-filter")) closeFilterMenus();
  });
}

async function refresh() {
  const db = await openDb();
  const entries = await Promise.all(slots.map(async (slot) => [slot.id, await getRecord(db, slot.id)]));
  db.close();
  state.records = new Map(entries);
  renderShell();
  await refreshDemandDashboard();
}

async function savePendingFile(slotId, file) {
  if (!file) return;
  await deleteSlot(slotId);
  const invalidMessage = getInvalidFileMessage(file);
  if (invalidMessage) {
    setLibraryState(invalidMessage);
    window.alert(invalidMessage);
    return;
  }

  const now = new Date().toISOString();
  const existing = { id: slotId };
  const record = {
    ...existing,
    id: slotId,
    applied: true,
    appliedAt: now,
    file,
    name: file.name,
    size: file.size,
    typeLabel: getFileTypeLabel(file),
    refreshMonth: getRefreshMonth(file.name, now),
    savedAt: now,
    pendingFile: file,
    pendingName: file.name,
    pendingSize: file.size,
    pendingTypeLabel: getFileTypeLabel(file),
    pendingRefreshMonth: getRefreshMonth(file.name, now),
    pendingSavedAt: now,
  };

  const db = await openDb();
  await putRecord(db, clearPendingFields(record));
  db.close();
  await refresh();
}

function getInvalidFileMessage(file) {
  return isSupportedTableFile(file?.name) ? "" : "请上传 .xlsx、.xls、.xlsm 或 .csv 表格文件。";
}

async function applySlot(slotId, options = {}) {
  const record = state.records.get(slotId);
  if (!record) return false;

  const targetName = record.pendingName || record.name;
  const invalidMessage = getInvalidFileMessage({ name: targetName });
  if (invalidMessage) {
    setLibraryState(invalidMessage);
    window.alert(invalidMessage);
    return false;
  }

  const appliedAt = new Date().toISOString();
  const updatedRecord = record.pendingFile
    ? clearPendingFields({
        ...record,
        file: record.pendingFile,
        name: record.pendingName,
        size: record.pendingSize,
        typeLabel: record.pendingTypeLabel,
        refreshMonth: record.pendingRefreshMonth,
        savedAt: record.pendingSavedAt,
        applied: true,
        appliedAt,
      })
    : {
        ...record,
        applied: true,
        appliedAt,
      };

  const db = await openDb();
  await putRecord(db, updatedRecord);
  db.close();
  if (!options.skipRefresh) await refresh();
  return true;
}

async function applyAllSlots() {
  const targetSlotIds = slots
    .filter((slot) => {
      const record = state.records.get(slot.id);
      return record?.pendingFile || record?.file;
    })
    .map((slot) => slot.id);

  if (els.applyAllButton) els.applyAllButton.disabled = true;
  setLibraryState("刷新应用中");
  try {
    for (const slotId of targetSlotIds) {
      const applied = await applySlot(slotId, { skipRefresh: true });
      if (applied === false) throw new Error("应用刷新失败");
    }
    await refresh();
    setLibraryState(targetSlotIds.length ? "已强制刷新应用" : "暂无可应用文件");
  } catch (error) {
    console.warn("apply all failed", error);
    setLibraryState(error.message || "刷新应用失败");
  } finally {
    updateApplyAllButton();
  }
}

async function deleteSlot(slotId) {
  const db = await openDb();
  await deleteRecord(db, slotId);
  db.close();
  await refresh();
}

async function clearLibraryCache() {
  const confirmed = window.confirm("确认清空本次上传的文件吗？清空后需要重新上传。");
  if (!confirmed) return;

  if (els.clearCacheButton) els.clearCacheButton.disabled = true;
  setLibraryState("清除中");
  try {
    await deleteLibraryDatabase();
    state.records = new Map();
    state.demandRows = [];
    state.filteredRows = [];
    resetAllFilterState();
    renderShell();
    setLibraryState("已清空本次文件");
  } catch (error) {
    console.warn("clear library cache failed", error);
    setLibraryState("清除失败");
    window.alert(error.message || "清空本次文件失败，请重试。");
  } finally {
    if (els.clearCacheButton) els.clearCacheButton.disabled = false;
  }
}

let dashboardRevision = 0;
async function refreshDemandDashboard() {
  const revision = ++dashboardRevision;
  state.demandRows = [];
  state.filteredRows = [];
  renderDemandDashboard("正在读取本次文件…");
  state.currentPage = 1;
  if (!window.XLSX) {
    state.demandRows = [];
    state.filteredRows = [];
    renderDemandDashboard("Excel 解析组件未加载，请刷新页面后重试");
    return;
  }

  const records = Object.fromEntries(
    Object.entries(SLOT_IDS).map(([key, slotId]) => {
      const record = state.records.get(slotId);
      return [key, record?.applied && record?.file && isSupportedTableFile(record.name) ? record : null];
    })
  );

  if (Object.values(records).some(record => !record)) {
    state.demandRows = [];
    state.filteredRows = [];
    resetAllFilterState();
    renderDemandDashboard("请上传本次最新的全部 4 份文件后生成明细");
    return;
  }

  try {
    const rows = await buildDemandAllocationRows(records);
    if (revision !== dashboardRevision) return;
    state.demandRows = rows;
  } catch (error) {
    if (revision !== dashboardRevision) return;
    console.warn("build demand allocation failed", error);
    state.demandRows = [];
    resetAllFilterState();
    renderDemandDashboard(error.message || "文件读取失败");
    setLibraryState(error.message || "文件读取失败");
    return;
  }
  rebuildFilterOptions();
  pruneFilterSelections();
  renderDemandDashboard();
}

async function buildDemandAllocationRows(records) {
  const [categoryRows, divisionRows, rosterRows, demandRows] = await Promise.all([
    readTableRows(records.category, "Dim-YL医疗器械商品分类"),
    readTableRows(records.purchaseDivision, "产品线明细", true),
    readTableRows(records.roster),
    readTableRows(records.demand),
  ]);
  const categoryIndex = buildCategoryIndex(categoryRows);
  const divisionIndex = buildPurchaseDivisionIndex(divisionRows);
  const minimumOrderIndex = buildMinimumOrderIndex(divisionRows);
  const rosterIndex = buildRosterIndex(rosterRows);

  const detailRows = demandRows.flatMap(({ rows, source }) => {
    const headerIndex = findHeaderRowIndex(rows, [
      ...TABLE_FIELD_ALIASES.demandApplicant,
      ...TABLE_FIELD_ALIASES.demandMaterialCode,
      ...TABLE_FIELD_ALIASES.demandProcessNo,
      ...TABLE_FIELD_ALIASES.demandQuantity,
    ]);
    if (headerIndex < 0) return [];
    const columnMap = buildColumnMapByAliases(rows[headerIndex], {
      applicant: TABLE_FIELD_ALIASES.demandApplicant,
      materialCode: TABLE_FIELD_ALIASES.demandMaterialCode,
      oaProcessNo: TABLE_FIELD_ALIASES.demandProcessNo,
      stockReason: TABLE_FIELD_ALIASES.demandStockReason,
      purchaseEntity: TABLE_FIELD_ALIASES.demandPurchaseEntity,
      requiredReadyDate: TABLE_FIELD_ALIASES.demandRequiredReadyDate,
    });
    columnMap.quantity = findDemandQuantityColumnIndex(rows[headerIndex], columnMap);
    if (columnMap.materialCode === undefined) return [];

    return rows.slice(headerIndex + 1).map((row, rowOffset) => {
      const applicant = getCellValue(row, columnMap.applicant);
      const materialCode = getCellValue(row, columnMap.materialCode);
      const quantityNumber = getDemandRowQuantity(row, columnMap);
      const quantity = Number.isFinite(quantityNumber) ? formatPlainNumber(quantityNumber) : "";
      const oaProcessNo = getCellValue(row, columnMap.oaProcessNo);
      const stockReason = getCellValue(row, columnMap.stockReason);
      const purchaseEntity = getCellValue(row, columnMap.purchaseEntity);
      const requiredReadyDate = getCellValue(row, columnMap.requiredReadyDate);
      if (!materialCode && !applicant && !quantity && !oaProcessNo) return null;

      const materialKey = normalizeLookupKey(materialCode);
      const category = categoryIndex.byMaterial.get(materialKey) || {};
      const purchaseGroup = category.purchaseGroup || "";
      const groupMaterialKey = makeGroupMaterialKey(purchaseGroup, materialCode);
      const groupDivisionRows = divisionIndex.byGroupMaterial.get(groupMaterialKey) || [];
      const materialDivisionRows = divisionIndex.byMaterial.get(materialKey) || [];
      const minimumOrderQuantityNumber = minimumOrderIndex.byMaterial.get(materialKey);
      const minimumOrderQuantity = Number.isFinite(minimumOrderQuantityNumber) ? formatPlainNumber(minimumOrderQuantityNumber) : "";
      const minimumOrderStatus = getMinimumOrderStatus(quantityNumber, minimumOrderQuantityNumber);
      const buyerRows = groupDivisionRows.length ? groupDivisionRows : materialDivisionRows;
      const buyer = joinUnique(buyerRows.map((item) => item.buyer));
      const supplierShortName = joinUnique(materialDivisionRows.map((item) => item.supplierShortName));

      return {
        businessUnit: rosterIndex.byName.get(normalizeLookupKey(applicant)) || "",
        productLine: category.productLine || "",
        salesSeries: category.salesSeries || "",
        purchaseGroup,
        buyer,
        applicant,
        supplierShortName,
        oaProcessNo,
        stockReason,
        purchaseEntity,
        materialCode,
        sku: category.sku || "",
        materialCodeValid: divisionIndex.materialCodes.has(materialKey) ? "是" : "否",
        materialName: category.materialName || "",
        quantity,
        requiredReadyDate,
        minimumOrderQuantity,
        minimumOrderStatus,
        quantityNumber,
        sourceFile: source.fileName,
        sourceSheet: source.sheetName,
        sourceSlotId: source.slotId,
        sourceRowNumber: headerIndex + rowOffset + 2,
      };
    }).filter(Boolean);
  });
  return mergeDemandRowsByMaterial(detailRows);
}

function mergeDemandRowsByMaterial(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    // 无物料编码的记录不能视为同一物料。
    const key = normalizeLookupKey(row.materialCode) || Symbol();
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });
  return [...groups.values()].map((items) => {
    const merged = { ...items[0] };
    const fields = new Set([
      ...DETAIL_COLUMNS.map((column) => column.key),
      ...FILTER_DEFINITIONS.map((filter) => filter.key),
    ]);
    fields.forEach((key) => {
      if (["materialCode", "quantity", "minimumOrderStatus"].includes(key)) return;
      const values = [...new Set(items.map((row) => String(row[key] ?? "").trim()).filter(Boolean))];
      merged[key] = values.join(key === "oaProcessNo" ? "/" : "、");
    });
    // 任一来源数量缺失时，不把部分合计当成完整需求量。
    const sum = items.every((row) => Number.isFinite(row.quantityNumber))
      ? items.reduce((total, row) => total + row.quantityNumber, 0)
      : NaN;
    merged.quantityNumber = Number.isFinite(sum) ? sum : NaN;
    merged.quantity = formatPlainNumber(merged.quantityNumber);
    const minimum = merged.minimumOrderQuantity === "" ? NaN : Number(merged.minimumOrderQuantity);
    merged.minimumOrderStatus = getMinimumOrderStatus(merged.quantityNumber, minimum);
    merged.sourceRows = items.map((row) => ({
      sourceFile: row.sourceFile,
      sourceSheet: row.sourceSheet,
      sourceSlotId: row.sourceSlotId,
      sourceRowNumber: row.sourceRowNumber,
      oaProcessNo: row.oaProcessNo,
      quantityNumber: row.quantityNumber,
    }));
    return merged;
  });
}

async function readTableRows(record, preferredSheetName = "", onlyExactSheet = false) {
  if (!record) return [];
  const workbook = await readWorkbook(record.file);
  if (onlyExactSheet && !workbook.SheetNames.includes(preferredSheetName)) {
    throw new Error(`${record.name}：缺少“${preferredSheetName}”工作表，请上传正确文件；不会读取其他 sheet`);
  }
  const sheetNames = onlyExactSheet ? [preferredSheetName] : getOrderedSheetNames(workbook, preferredSheetName);
  return sheetNames.map((sheetName) => ({
    rows: sheetToRows(workbook.Sheets[sheetName]),
    source: {
      fileName: record.name,
      slotId: record.id,
      sheetName,
    },
  })).filter((table) => table.rows.length);
}

function getOrderedSheetNames(workbook, preferredSheetName) {
  const sheetNames = workbook.SheetNames || [];
  if (!preferredSheetName) return sheetNames;
  const preferred = sheetNames.find((sheetName) => sheetName === preferredSheetName)
    || sheetNames.find((sheetName) => sheetName.includes(preferredSheetName));
  if (!preferred) return sheetNames;
  return [preferred, ...sheetNames.filter((sheetName) => sheetName !== preferred)];
}

async function readWorkbook(file) {
  const extension = String(file?.name || "").split(".").pop()?.toLowerCase();
  if (extension === "csv") {
    return window.XLSX.read(await file.text(), {
      type: "string",
      raw: false,
      cellText: true,
    });
  }
  return window.XLSX.read(await file.arrayBuffer(), {
    type: "array",
    cellText: true,
    cellStyles: true,
  });
}

function sheetToRows(sheet) {
  if (!sheet) return [];
  return window.XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    blankrows: false,
    raw: false,
  });
}

function buildCategoryIndex(tables) {
  const byMaterial = new Map();
  tables.forEach(({ rows }) => {
    const headerIndex = findHeaderRowIndex(rows, ["物料编码", "识别码", "SKU", "销售产品线", "销售系列", "采购分组", "采购组", "金蝶名称"]);
    const columnMap = headerIndex >= 0
      ? buildColumnMapByAliases(rows[headerIndex], {
        materialCode: ["物料编码", "识别码", "物料编号", "商品编码", "存货编码", "产品编码"],
        sku: ["SKU", "sku", "Sku"],
        productLine: TABLE_FIELD_ALIASES.productLine,
        salesSeries: TABLE_FIELD_ALIASES.salesSeries,
        purchaseGroup: TABLE_FIELD_ALIASES.purchaseGroup,
        materialName: ["金蝶名称", "物料名称", "物料名", "商品名称", "品名", "产品名称"],
      })
      : {};
    const startIndex = headerIndex >= 0 ? headerIndex + 1 : 0;
    rows.slice(startIndex).forEach((row) => {
      const materialCode = getCellValue(row, 0) || getCellValue(row, columnMap.materialCode);
      const materialKey = normalizeLookupKey(materialCode);
      if (!materialKey) return;
      byMaterial.set(materialKey, {
        sku: getCellValue(row, 2) || getCellValue(row, columnMap.sku),
        productLine: getCellValue(row, columnMap.productLine ?? 6),
        salesSeries: getCellValue(row, columnMap.salesSeries ?? 7),
        purchaseGroup: getCellValue(row, 21) || getCellValue(row, columnMap.purchaseGroup),
        materialName: getCellValue(row, columnMap.materialName ?? 3),
      });
    });
  });
  return { byMaterial };
}

function buildPurchaseDivisionIndex(tables) {
  const byMaterial = new Map();
  const byGroupMaterial = new Map();
  const materialCodes = new Set();
  tables.forEach(({ rows }) => {
    const headerIndex = findHeaderRowIndex(rows, [
      ...TABLE_FIELD_ALIASES.divisionPurchaseGroup,
      ...TABLE_FIELD_ALIASES.divisionMaterialCode,
      ...TABLE_FIELD_ALIASES.divisionBuyer,
      ...TABLE_FIELD_ALIASES.divisionSupplier,
    ]);
    if (headerIndex < 0) return;
    const columnMap = buildColumnMapByAliases(rows[headerIndex], {
      purchaseGroup: TABLE_FIELD_ALIASES.divisionPurchaseGroup,
      materialCode: TABLE_FIELD_ALIASES.divisionMaterialCode,
      buyer: TABLE_FIELD_ALIASES.divisionBuyer,
      supplierShortName: TABLE_FIELD_ALIASES.divisionSupplier,
    });
    if (columnMap.materialCode === undefined) return;

    rows.slice(headerIndex + 1).forEach((row) => {
      const materialCode = getCellValue(row, columnMap.materialCode);
      const materialKey = normalizeLookupKey(materialCode);
      if (!materialKey) return;
      const item = {
        purchaseGroup: getCellValue(row, columnMap.purchaseGroup),
        materialCode,
        buyer: getCellValue(row, columnMap.buyer),
        supplierShortName: getCellValue(row, columnMap.supplierShortName),
      };
      materialCodes.add(materialKey);
      pushMapArray(byMaterial, materialKey, item);
      pushMapArray(byGroupMaterial, makeGroupMaterialKey(item.purchaseGroup, materialCode), item);
    });
  });
  return { byMaterial, byGroupMaterial, materialCodes };
}

function buildMinimumOrderIndex(tables) {
  const byMaterial = new Map();
  tables
    .filter((table) => String(table.source?.sheetName || "").trim().includes("产品线明细"))
    .forEach(({ rows }) => {
      const headerIndex = findHeaderRowIndex(rows, TABLE_FIELD_ALIASES.divisionMaterialCode, 1);
      const columnMap = headerIndex >= 0
        ? buildColumnMapByAliases(rows[headerIndex], {
          materialCode: TABLE_FIELD_ALIASES.divisionMaterialCode,
        })
        : {};
      const startIndex = headerIndex >= 0 ? headerIndex + 1 : 0;
      rows.slice(startIndex).forEach((row) => {
        const materialCode = getCellValue(row, columnMap.materialCode ?? 0);
        const materialKey = normalizeLookupKey(materialCode);
        if (!materialKey) return;
        const minimumOrderQuantity = forceDemandNumber(row?.[9]);
        if (Number.isFinite(minimumOrderQuantity)) byMaterial.set(materialKey, minimumOrderQuantity);
      });
    });
  return { byMaterial };
}

function getMinimumOrderStatus(quantityNumber, minimumOrderQuantityNumber) {
  if (!Number.isFinite(quantityNumber) || !Number.isFinite(minimumOrderQuantityNumber)) return "";
  return quantityNumber >= minimumOrderQuantityNumber ? "满足" : "不满足";
}

function buildRosterIndex(tables) {
  const byName = new Map();
  tables.forEach(({ rows }) => {
    rows.forEach((row) => {
      const name = getCellValue(row, 0);
      const businessUnit = getCellValue(row, 1);
      const nameKey = normalizeLookupKey(name);
      if (!nameKey || !businessUnit || normalizeDemandHeader(name) === "姓名") return;
      byName.set(nameKey, businessUnit);
    });
  });
  return { byName };
}

function findHeaderRowIndex(rows, aliases, minimumScore = 2) {
  let bestIndex = -1;
  let bestScore = 0;
  const normalizedAliases = aliases.map(normalizeDemandHeader).filter(Boolean);
  rows.slice(0, 30).forEach((row, index) => {
    const normalizedHeaders = row.map(normalizeDemandHeader);
    const score = normalizedHeaders.filter((header) =>
      header && normalizedAliases.some((alias) => header === alias || header.includes(alias) || alias.includes(header))
    ).length;
    if (score > bestScore) {
      bestIndex = index;
      bestScore = score;
    }
  });
  return bestScore >= minimumScore ? bestIndex : -1;
}

function buildColumnMapByAliases(headerRow = [], aliasMap = {}) {
  const normalizedHeaders = headerRow.map(normalizeDemandHeader);
  const columnMap = {};
  Object.entries(aliasMap).forEach(([key, aliases]) => {
    const normalizedAliases = aliases.map(normalizeDemandHeader).filter(Boolean);
    const exactIndex = normalizedHeaders.findIndex((header) => normalizedAliases.includes(header));
    if (exactIndex >= 0) {
      columnMap[key] = exactIndex;
      return;
    }

    const looseIndex = normalizedHeaders.findIndex((header) =>
      header && normalizedAliases.some((alias) => alias && (header.includes(alias) || alias.includes(header)))
    );
    if (looseIndex >= 0) columnMap[key] = looseIndex;
  });
  return columnMap;
}

function findExactHeaderColumnIndex(headerRow = [], headerName) {
  const normalizedHeaderName = normalizeDemandHeader(headerName);
  const index = headerRow.findIndex((header) => normalizeDemandHeader(header) === normalizedHeaderName);
  return index >= 0 ? index : undefined;
}

function normalizeDemandHeader(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[()[\]{}（）【】<>《》:：]/g, "")
    .toLowerCase();
}

function getCellValue(row, columnIndex) {
  if (columnIndex === undefined || columnIndex < 0) return "";
  return String(row?.[columnIndex] ?? "").trim();
}

function getDemandRowQuantity(row, columnMap) {
  const value = row?.[columnMap.quantity];
  if (typeof value === "number") return Number.isFinite(value) ? value : NaN;
  const text = String(value ?? "").trim();
  if (!/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(text)) return NaN;
  const number = Number(text.replace(/,/g, ""));
  return Number.isFinite(number) ? number : NaN;
}

function findDemandQuantityColumnIndex(headerRow, columnMap) {
  const quantityColumns = headerRow.flatMap((header, index) =>
    normalizeDemandHeader(header) === "数量" ? [index] : []
  );
  if (quantityColumns.length === 1) return quantityColumns[0];
  // OA 导出同时包含调库存数量和备货明细数量，后者紧邻货好时间。
  const detailQuantityIndex = columnMap.requiredReadyDate - 1;
  return quantityColumns.includes(detailQuantityIndex) ? detailQuantityIndex : undefined;
}

function normalizeLookupKey(value) {
  return String(value ?? "").trim().replace(/\s+/g, "").toLowerCase();
}

function makeGroupMaterialKey(purchaseGroup, materialCode) {
  return `${normalizeLookupKey(purchaseGroup)}|${normalizeLookupKey(materialCode)}`;
}

function pushMapArray(map, key, value) {
  if (!key) return;
  if (!map.has(key)) map.set(key, []);
  map.get(key).push(value);
}

function joinUnique(values) {
  return uniqueSortedValues(values).join("、");
}

function parseDemandQuantity(value) {
  return forceDemandNumber(value);
}

function forceDemandNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : NaN;
  const match = String(value ?? "").replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
  if (!match) return NaN;
  const number = Number(match[0]);
  return Number.isFinite(number) ? number : NaN;
}

function formatPlainNumber(value) {
  if (!Number.isFinite(value)) return "";
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(6)));
}

function isSupportedTableFile(fileName) {
  return /\.(xlsx|xlsm|xls|csv)$/i.test(String(fileName || ""));
}

function rebuildFilterOptions() {
  state.filterOptions = Object.fromEntries(
    FILTER_DEFINITIONS.map((filter) => [
      filter.key,
      uniqueSortedValues(getRowsForFilterOptions(filter.key).flatMap((row) => getRowFilterValues(row, filter.key))),
    ])
  );
}

function getRowsForFilterOptions(targetKey) {
  const linkedKeys = getLinkedFilterKeys(targetKey);
  if (!linkedKeys.length) return state.demandRows;
  return state.demandRows.filter((row) =>
    linkedKeys.every((key) => {
      const selected = state.filters.get(key);
      if (!selected?.size) return true;
      return getRowFilterValues(row, key).some((value) => selected.has(value));
    })
  );
}

function getLinkedFilterKeys(targetKey) {
  const group = LINKED_FILTER_GROUPS.find((keys) => keys.includes(targetKey));
  return group ? group.filter((key) => key !== targetKey) : [];
}

function uniqueSortedValues(values) {
  return [...new Set(values.map((value) => String(value ?? "").trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "zh-CN", { numeric: true }));
}

function pruneFilterSelections() {
  FILTER_DEFINITIONS.forEach((filter) => {
    const options = new Set(state.filterOptions[filter.key] || []);
    const selected = state.filters.get(filter.key) || new Set();
    state.filters.set(filter.key, new Set([...selected].filter((value) => options.has(value))));
  });
}

function resetAllFilterState() {
  FILTER_DEFINITIONS.forEach((filter) => state.filters.set(filter.key, new Set()));
  state.filterOptions = Object.fromEntries(FILTER_DEFINITIONS.map((filter) => [filter.key, []]));
}

function renderShell() {
  if (els.slotCount) els.slotCount.textContent = String(slots.length);
  if (els.uploadedCount) els.uploadedCount.textContent = String(countUploadedRecords());
  if (els.appliedCount) els.appliedCount.textContent = String(countAppliedRecords());
  if (els.sourceNote) els.sourceNote.textContent = `本次上传（不使用缓存）｜引用时间：${getLatestAppliedTime()}`;
  if (els.slotGrid) els.slotGrid.innerHTML = slots.map(renderSlot).join("");
  renderDemandDashboard();
  setLibraryState("仅引用本次上传");
  updateApplyAllButton();
}

function renderSlot(slot) {
  const record = state.records.get(slot.id);
  const display = getDisplayRecord(record);
  const hasPending = Boolean(record?.pendingFile);
  const isApplied = Boolean(record?.applied && !hasPending);
  return `
    <article class="slot-card ${isApplied ? "applied" : ""}" data-drop="${slot.id}">
      <div class="slot-head">
        <div>
          <p class="eyebrow">Slot</p>
          <h2>${escapeHtml(slot.label)}</h2>
        </div>
        <span class="slot-status ${isApplied ? "applied" : hasPending ? "pending" : ""}">
          ${isApplied ? "已引用" : hasPending ? "待应用" : "未上传"}
        </span>
      </div>
      <div class="file-meta">
        <span>${escapeHtml(display?.name || "未上传文件")}</span>
        <strong>${display ? `${escapeHtml(display.typeLabel)} / ${formatFileSize(display.size)}` : "--"}</strong>
        <small>刷新月份：${escapeHtml(display?.refreshMonth || "--")}</small>
        <small>更新时间：${formatDateTime(display?.savedAt)}</small>
        <small>引用时间：${formatDateTime(record?.appliedAt)}</small>
      </div>
      <div class="slot-actions">
        <label class="upload-button">
          <input type="file" accept="${TABLE_FILE_ACCEPT}" data-upload="${slot.id}" />
          上传/替换
        </label>
        <button type="button" data-apply="${slot.id}" ${record?.pendingFile || record?.file ? "" : "disabled"}>应用刷新</button>
        <button class="danger-button" type="button" data-delete="${slot.id}" ${record ? "" : "disabled"}>删除</button>
      </div>
    </article>
  `;
}

function renderDemandDashboard(emptyMessage = "暂无匹配数据") {
  state.filteredRows = getFilteredDemandRows();
  clampCurrentPage();
  updateDemandMetrics();
  renderFilterControls();
  renderDemandTable(emptyMessage);
  renderPagination();
}

function getFilteredDemandRows() {
  return state.demandRows.filter((row) =>
    FILTER_DEFINITIONS.every((filter) => {
      const selected = state.filters.get(filter.key);
      if (!selected?.size) return true;
      return getRowFilterValues(row, filter.key).some((value) => selected.has(value));
    })
  );
}

function updateDemandMetrics() {
  const quantityValues = state.filteredRows.map((row) => row.quantityNumber).filter(Number.isFinite);
  const total = quantityValues.reduce((sum, value) => sum + value, 0);
  const unresolvedCount = state.filteredRows.length - quantityValues.length;
  const quantityWarning = unresolvedCount ? `（${unresolvedCount} 条数量待核对）` : "";
  const supplierCount = new Set(state.filteredRows.flatMap((row) => splitMultiValue(row.supplierShortName))).size;
  if (els.beihuoTotal) els.beihuoTotal.textContent = unresolvedCount ? "待核对" : formatMetricNumber(total);
  if (els.quantityTotal) els.quantityTotal.textContent = `数量合计：${formatMetricNumber(total)}${quantityWarning}`;
  if (els.supplierCount) els.supplierCount.textContent = formatMetricNumber(supplierCount);
  if (els.downloadDetailButton) els.downloadDetailButton.disabled = !state.filteredRows.length || !window.XLSX;
}

function renderFilterControls() {
  FILTER_DEFINITIONS.forEach((filter) => {
    const container = els.filterToolbar?.querySelector(`[data-filter="${filter.key}"]`);
    if (!container) return;
    const options = state.filterOptions[filter.key] || [];
    const selected = state.filters.get(filter.key) || new Set();
    container.innerHTML = `
      <button class="multi-filter-button" type="button" data-filter-button="${filter.key}" ${options.length ? "" : "disabled"}>
        <span>${escapeHtml(getFilterButtonLabel(filter, selected))}</span>
        <i aria-hidden="true">▾</i>
      </button>
      <div class="multi-filter-menu">
        <label class="multi-filter-option">
          <input type="checkbox" data-filter-option data-filter-key="${filter.key}" value="__all__" ${selected.size ? "" : "checked"} />
          <span>全部</span>
        </label>
        ${
          options.length
            ? options.map((value) => renderFilterOption(filter.key, value, selected.has(value))).join("")
            : `<div class="multi-filter-empty">暂无选项</div>`
        }
      </div>
    `;
  });
}

function renderFilterOption(filterKey, value, checked) {
  const safeValue = escapeHtml(value);
  return `
    <label class="multi-filter-option ${checked ? "selected" : ""}">
      <input type="checkbox" data-filter-option data-filter-key="${filterKey}" value="${safeValue}" ${checked ? "checked" : ""} />
      <span>${safeValue}</span>
    </label>
  `;
}

function getFilterButtonLabel(filter, selected) {
  if (!selected?.size) return `全部${filter.label}`;
  const values = [...selected];
  if (values.length === 1) return values[0];
  if (values.length === 2) return values.join("、");
  return `已选${values.length}项`;
}

function getRowFilterValues(row, key) {
  return splitMultiValue(row[key]);
}

function splitMultiValue(value) {
  return String(value ?? "")
    .split(/[、,，;；/]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function handleFilterToolbarClick(event) {
  const button = event.target.closest("[data-filter-button]");
  if (!button) return;
  event.stopPropagation();
  const container = button.closest(".multi-filter");
  const shouldOpen = !container.classList.contains("open");
  closeFilterMenus();
  if (shouldOpen) container.classList.add("open");
}

function handleFilterToolbarChange(event) {
  const input = event.target.closest("[data-filter-option]");
  if (!input) return;
  const selected = new Set(state.filters.get(input.dataset.filterKey) || []);
  if (input.value === "__all__") {
    selected.clear();
  } else if (input.checked) {
    selected.add(input.value);
  } else {
    selected.delete(input.value);
  }
  state.filters.set(input.dataset.filterKey, selected);
  state.currentPage = 1;
  rebuildFilterOptions();
  pruneFilterSelections();
  renderDemandDashboard();
}

function clearAllFilters() {
  FILTER_DEFINITIONS.forEach((filter) => state.filters.set(filter.key, new Set()));
  state.currentPage = 1;
  closeFilterMenus();
  rebuildFilterOptions();
  renderDemandDashboard();
}

function closeFilterMenus() {
  els.filterToolbar?.querySelectorAll(".multi-filter.open").forEach((container) => container.classList.remove("open"));
}

function renderDemandTable(emptyMessage) {
  if (!els.detailTableBody || !els.detailEmpty) return;
  const rows = getCurrentPageRows();
  els.detailTableBody.innerHTML = rows.map(renderDemandTableRow).join("");
  els.detailEmpty.hidden = state.filteredRows.length > 0;
  els.detailEmpty.textContent = state.demandRows.length ? emptyMessage : "请上传本次最新的全部 4 份文件后生成明细";
  if (els.detailState) {
    const startRow = (state.currentPage - 1) * DETAIL_PAGE_SIZE + 1;
    const endRow = Math.min(state.currentPage * DETAIL_PAGE_SIZE, state.filteredRows.length);
    els.detailState.textContent = state.filteredRows.length
      ? `第 ${formatMetricNumber(startRow)}–${formatMetricNumber(endRow)} 条，共 ${formatMetricNumber(state.filteredRows.length)} 条`
      : state.demandRows.length ? "无匹配数据" : "暂无数据";
  }
}

function getPageCount() {
  return Math.max(1, Math.ceil(state.filteredRows.length / DETAIL_PAGE_SIZE));
}

function clampCurrentPage() {
  state.currentPage = Math.min(Math.max(1, state.currentPage), getPageCount());
}

function getCurrentPageRows() {
  const startIndex = (state.currentPage - 1) * DETAIL_PAGE_SIZE;
  return state.filteredRows.slice(startIndex, startIndex + DETAIL_PAGE_SIZE);
}

function changeDetailPage(offset) {
  const nextPage = Math.min(Math.max(1, state.currentPage + offset), getPageCount());
  if (nextPage === state.currentPage) return;
  state.currentPage = nextPage;
  renderDemandTable();
  renderPagination();
  els.detailTableBody?.closest(".table-wrap")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderPagination() {
  if (!els.paginationControls) return;
  const totalRows = state.filteredRows.length;
  const pageCount = getPageCount();
  els.paginationControls.hidden = totalRows === 0;
  if (els.paginationInfo) {
    els.paginationInfo.textContent = `第 ${state.currentPage} / ${pageCount} 页｜共 ${formatMetricNumber(totalRows)} 条｜每页 ${DETAIL_PAGE_SIZE} 条`;
  }
  if (els.previousPageButton) els.previousPageButton.disabled = state.currentPage <= 1;
  if (els.nextPageButton) els.nextPageButton.disabled = state.currentPage >= pageCount;
}

function renderDemandTableRow(row) {
  return `
    <tr>
      ${DETAIL_COLUMNS.map((column) => `<td>${escapeHtml(row[column.key] || "")}</td>`).join("")}
    </tr>
  `;
}

async function downloadDetailWorkbook() {
  if (!state.filteredRows.length) {
    window.alert("暂无可下载的明细数据。");
    return;
  }
  if (!window.XLSX) {
    window.alert("Excel 导出组件未加载，请刷新页面后重试。");
    return;
  }

  const workbook = window.XLSX.utils.book_new();
  const sheetRows = [
    DETAIL_COLUMNS.map((column) => column.label),
    ...state.filteredRows.map((row) => DETAIL_COLUMNS.map((column) => {
      if (column.key !== "quantity") return row[column.key] || "";
      const quantity = Number.isFinite(row.quantityNumber)
        ? row.quantityNumber : getDemandRowQuantity([row.quantity], { quantity: 0 });
      return Number.isFinite(quantity) ? quantity : "";
    })),
  ];
  const worksheet = window.XLSX.utils.aoa_to_sheet(sheetRows);
  const quantityColumn = DETAIL_COLUMNS.findIndex((column) => column.key === "quantity");
  for (let rowIndex = 1; rowIndex < sheetRows.length; rowIndex += 1) {
    const cell = worksheet[window.XLSX.utils.encode_cell({ r: rowIndex, c: quantityColumn })];
    if (cell?.t === "n") cell.z = "0";
  }
  worksheet["!cols"] = getAutoExportColumns(sheetRows);
  window.XLSX.utils.book_append_sheet(workbook, worksheet, "备货需求分配明细");
  const sheetShapes = [{ rowCount: sheetRows.length, columnCount: DETAIL_COLUMNS.length }];

  const fileName = `备货需求分配明细_${formatFileTimestamp(new Date())}.xlsx`;
  if (!window.JSZip) {
    window.XLSX.writeFile(workbook, fileName);
    return;
  }

  try {
    await writeStyledDetailWorkbook(workbook, sheetShapes, fileName);
  } catch (error) {
    console.warn("styled workbook export failed", error);
    window.XLSX.writeFile(workbook, fileName);
  }
}

async function writeStyledDetailWorkbook(workbook, sheetShapes, fileName) {
  const workbookArray = window.XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });
  const zip = await window.JSZip.loadAsync(workbookArray);
  zip.file("xl/styles.xml", buildDetailWorkbookStylesXml());

  await Promise.all(sheetShapes.map(async (shape, index) => {
    const path = `xl/worksheets/sheet${index + 1}.xml`;
    const file = zip.file(path);
    if (!file) return;
    const xml = await file.async("string");
    zip.file(path, styleDetailWorksheetXml(xml, shape));
  }));

  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  downloadBlob(blob, fileName);
}

function buildDetailWorkbookStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFB7C9D6"/></left>
      <right style="thin"><color rgb="FFB7C9D6"/></right>
      <top style="thin"><color rgb="FFB7C9D6"/></top>
      <bottom style="thin"><color rgb="FFB7C9D6"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="0"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="0"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleMedium9"/>
</styleSheet>`;
}

function styleDetailWorksheetXml(xml, shape) {
  const parser = new DOMParser();
  const documentXml = parser.parseFromString(xml, "application/xml");
  const parseError = documentXml.querySelector("parsererror");
  if (parseError) return xml;

  const sheetData = documentXml.querySelector("sheetData");
  if (!sheetData) return xml;

  const wrapColumnIndex = DETAIL_COLUMNS.findIndex((column) => column.key === "stockReason") + 1;
  const quantityColumnIndex = DETAIL_COLUMNS.findIndex((column) => column.key === "quantity") + 1;
  for (let rowIndex = 1; rowIndex <= shape.rowCount; rowIndex += 1) {
    const row = ensureWorksheetRow(documentXml, sheetData, rowIndex);
    for (let columnIndex = 1; columnIndex <= shape.columnCount; columnIndex += 1) {
      const cell = ensureWorksheetCell(documentXml, row, rowIndex, columnIndex);
      const styleId = rowIndex === 1 ? "1" : columnIndex === quantityColumnIndex ? "4" : columnIndex === wrapColumnIndex ? "3" : "2";
      cell.setAttribute("s", styleId);
    }
  }

  return new XMLSerializer().serializeToString(documentXml);
}

function ensureWorksheetRow(documentXml, sheetData, rowIndex) {
  let row = sheetData.querySelector(`row[r="${rowIndex}"]`);
  if (row) return row;

  row = documentXml.createElementNS(sheetData.namespaceURI, "row");
  row.setAttribute("r", String(rowIndex));
  const nextRow = [...sheetData.children].find((item) => Number(item.getAttribute("r")) > rowIndex);
  sheetData.insertBefore(row, nextRow || null);
  return row;
}

function ensureWorksheetCell(documentXml, row, rowIndex, columnIndex) {
  const cellRef = `${columnIndexToName(columnIndex)}${rowIndex}`;
  let cell = row.querySelector(`c[r="${cellRef}"]`);
  if (cell) return cell;

  cell = documentXml.createElementNS(row.namespaceURI, "c");
  cell.setAttribute("r", cellRef);
  const nextCell = [...row.children].find((item) => cellNameToColumnIndex(item.getAttribute("r")) > columnIndex);
  row.insertBefore(cell, nextCell || null);
  return cell;
}

function columnIndexToName(index) {
  let name = "";
  let current = index;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    current = Math.floor((current - 1) / 26);
  }
  return name;
}

function cellNameToColumnIndex(cellName) {
  const letters = String(cellName || "").match(/[A-Z]+/i)?.[0] || "";
  return letters.toUpperCase().split("").reduce((sum, letter) => sum * 26 + letter.charCodeAt(0) - 64, 0);
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

function groupRowsByBuyer(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    const buyer = String(row.buyer || "").trim() || "未填写下单人";
    if (!groups.has(buyer)) groups.set(buyer, []);
    groups.get(buyer).push(row);
  });
  return [...groups.entries()].sort(([left], [right]) => left.localeCompare(right, "zh-CN", { numeric: true }));
}

function makeUniqueSheetName(rawName, usedSheetNames) {
  const baseName = sanitizeSheetName(rawName) || "未填写下单人";
  let sheetName = baseName.slice(0, 31);
  let index = 2;
  while (usedSheetNames.has(sheetName)) {
    const suffix = `_${index}`;
    sheetName = `${baseName.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }
  usedSheetNames.add(sheetName);
  return sheetName;
}

function sanitizeSheetName(value) {
  return String(value || "")
    .replace(/[\[\]:*?/\\]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function getExportColumnWidth(key) {
  return {
    purchaseGroup: 14,
    buyer: 18,
    businessUnit: 16,
    applicant: 14,
    supplierShortName: 20,
    oaProcessNo: 20,
    stockReason: 90,
    purchaseEntity: 18,
    materialCode: 18,
    sku: 16,
    materialCodeValid: 18,
    materialName: 28,
    quantity: 12,
    requiredReadyDate: 18,
    minimumOrderQuantity: 12,
    minimumOrderStatus: 14,
  }[key] || 14;
}

function getExportColumnMaxWidth(key) {
  return {
    stockReason: 120,
  }[key] || 255;
}

function getAutoExportColumns(sheetRows) {
  return DETAIL_COLUMNS.map((column, columnIndex) => {
    const contentWidth = sheetRows.reduce((width, row) => {
      const cellWidth = getTextDisplayWidth(row?.[columnIndex]);
      return Math.max(width, cellWidth);
    }, 0);
    const minWidth = getExportColumnWidth(column.key);
    const maxWidth = getExportColumnMaxWidth(column.key);
    return { wch: Math.min(Math.max(contentWidth + 2, minWidth), maxWidth) };
  });
}

function getTextDisplayWidth(value) {
  return String(value ?? "")
    .split(/\r?\n/)
    .reduce((maxWidth, line) => {
      const lineWidth = [...line].reduce((width, char) => width + (char.charCodeAt(0) > 255 ? 2 : 1), 0);
      return Math.max(maxWidth, lineWidth);
    }, 0);
}

function getDisplayRecord(record) {
  if (!record) return null;
  if (record.pendingFile) {
    return {
      name: record.pendingName,
      size: record.pendingSize,
      typeLabel: record.pendingTypeLabel,
      refreshMonth: record.pendingRefreshMonth,
      savedAt: record.pendingSavedAt,
    };
  }
  return record;
}

function countUploadedRecords() {
  return slots.filter((slot) => {
    const record = state.records.get(slot.id);
    return record?.file || record?.pendingFile;
  }).length;
}

function countAppliedRecords() {
  return slots.filter((slot) => state.records.get(slot.id)?.applied).length;
}

function updateApplyAllButton() {
  if (els.applyAllButton) els.applyAllButton.disabled = false;
}

function setLibraryState(text) {
  if (els.libraryState) els.libraryState.textContent = text;
}

function clearPendingFields(record) {
  const nextRecord = { ...record };
  delete nextRecord.pendingFile;
  delete nextRecord.pendingName;
  delete nextRecord.pendingSize;
  delete nextRecord.pendingTypeLabel;
  delete nextRecord.pendingRefreshMonth;
  delete nextRecord.pendingSavedAt;
  return nextRecord;
}

function getRefreshMonth(fileName, fallbackTime) {
  const name = String(fileName || "");
  const compact = name.match(/(20\d{2})(0[1-9]|1[0-2])/);
  if (compact) return `${compact[1]}-${compact[2]}`;
  const separated = name.match(/(20\d{2})[-_.年]+(0?[1-9]|1[0-2])/);
  if (separated) return `${separated[1]}-${String(separated[2]).padStart(2, "0")}`;
  return getMonthFromDate(fallbackTime);
}

function getMonthFromDate(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function getLatestAppliedTime() {
  const times = slots
    .map((slot) => state.records.get(slot.id)?.appliedAt)
    .filter(Boolean)
    .sort();
  if (!times.length) return "--";
  return formatDateTime(times.at(-1));
}

function getFileTypeLabel(file) {
  const extension = String(file?.name || "").split(".").pop()?.toLowerCase();
  if (extension === "xlsx" || extension === "xls" || extension === "xlsm") return "Excel 工作簿";
  if (extension === "csv") return "CSV 文件";
  return file?.type || "表格文件";
}

function formatDateTime(value) {
  if (!value) return "--";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";
  return new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function formatFileSize(size = 0) {
  if (!Number.isFinite(size) || size <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let value = size;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(value >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

function formatFileTimestamp(date) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
    String(date.getHours()).padStart(2, "0"),
    String(date.getMinutes()).padStart(2, "0"),
  ].join("");
}

function formatMetricNumber(value) {
  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: Number.isInteger(value) ? 0 : 2,
  }).format(value || 0);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  })[char]);
}

// Keep storage helpers asynchronous so the existing upload/apply flow stays unchanged.
async function openDb() { return { close() {} }; }
async function getRecord(db, key) { return sessionFiles.get(key); }
async function putRecord(db, record) { sessionFiles.set(record.id, record); }
async function deleteRecord(db, key) { sessionFiles.delete(key); }
async function deleteLibraryDatabase() { sessionFiles.clear(); }

init().catch((error) => {
  console.error(error);
  setLibraryState("读取失败");
});
