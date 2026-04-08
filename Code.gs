const APP_CONFIG = Object.freeze({
  timezone: 'Asia/Bangkok',
  sheets: {
    form: 'ข้อมูลฟอร์มระบบ',
    legacyForm: 'กรุณากรอกข้อมูล',
    rateStatement: 'ใบแจ้งอัตราค่าบริการ',
    serviceInvoice: 'ใบแจ้งค่าบริการ',
    catalog: 'Test No.',
    ledger: 'ทะเบียนลูกหนี้',
    dashboard: 'Dashboard',
    receiptSummary: 'สรุปใบเสร็จ'
  },
  form: {
    maxRows: 20,
    rateRange: 'B3:D22',
    invoiceRange: 'F3:H22',
    invoiceNumberCell: 'K3',
    issueDateThaiCell: 'K4',
    issueDateEnglishCell: 'K5',
    customerNameCell: 'K6',
    customerAddressCell: 'K7',
    recordIdCell: 'M2',
    ledgerRowCell: 'M3',
    rawIssueDateCell: 'M4'
  },
  totalCells: {
    rateStatement: 'H38',
    serviceInvoice: 'H38'
  },
  pdfExport: {
    exportFormat: 'pdf',
    format: 'pdf',
    attachment: true,
    single: true,
    size: 'A4',
    portrait: true,
    fitw: true,
    gridlines: false,
    sheetnames: false,
    printtitle: false,
    pagenumbers: false,
    fzr: false,
    top_margin: 0.5,
    bottom_margin: 0.5,
    left_margin: 0.3,
    right_margin: 0.3
  },
  pdfRetry: {
    maxAttempts: 4,
    initialDelayMs: 900
  },
  ledger: {
    coreHeaders: [
      'ชื่อผู้ใช้บริการ',
      'ที่อยู่',
      'วันที่ออกใบแจ้งหนี้',
      'เลขที่ใบแจ้งอัตราค่าบริการ',
      'วันที่ออกเลขที่ใบแจ้งอัตราค่าบริการ',
      'จำนวนเงิน',
      'วันที่ชำระหนี้',
      'สถานะใบเสร็จ'
    ],
    extraHeaders: [
      'รหัสรายการ',
      'สถานะระบบ',
      'ข้อมูลรายการอัตราค่าบริการ (ระบบ)',
      'ข้อมูลรายการใบแจ้งค่าบริการ (ระบบ)',
      'PDF ใบแจ้งอัตราค่าบริการ',
      'PDF ใบแจ้งค่าบริการ',
      'ยอดใบแจ้งอัตราค่าบริการ',
      'ยอดใบแจ้งค่าบริการ',
      'สร้างเมื่อ',
      'แก้ไขเมื่อ',
      'ยกเลิกเมื่อ',
      'หมายเหตุยกเลิก',
      'ยอดค้างชำระ'
    ],
    systemStatus: {
      active: 'ACTIVE',
      cancelled: 'CANCELLED'
    },
    displayStatus: {
      pending: 'รอชำระ',
      paid: 'ชำระแล้ว',
      cancelled: 'ยกเลิกใบแจ้งหนี้'
    }
  },
  colors: {
    brand: '#0f766e',
    brandSoft: '#ccfbf1',
    pending: '#fff7ed',
    paid: '#ecfdf5',
    cancelled: '#fef2f2',
    pendingAccent: '#f97316',
    paidAccent: '#16a34a',
    cancelledAccent: '#dc2626',
    header: '#0f172a',
    muted: '#e2e8f0'
  }
});

const LEDGER_COLUMNS = Object.freeze({
  customerName: 1,
  address: 2,
  invoiceDate: 3,
  rateNumber: 4,
  rateDate: 5,
  amount: 6,
  paidDate: 7,
  receiptStatus: 8,
  recordId: 9,
  systemStatus: 10,
  rateItemsJson: 11,
  invoiceItemsJson: 12,
  ratePdfUrl: 13,
  invoicePdfUrl: 14,
  rateTotal: 15,
  invoiceTotal: 16,
  createdAt: 17,
  updatedAt: 18,
  cancelledAt: 19,
  cancelReason: 20,
  outstandingAmount: 21,
  lastColumn: 21
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ระบบลูกหนี้')
    .addItem('ติดตั้ง / ตั้งค่าระบบ', 'setupSystem')
    .addItem('เริ่มใบแจ้งหนี้ใหม่', 'prepareNewInvoice')
    .addSeparator()
    .addItem('บันทึกข้อมูลจากฟอร์มระบบ', 'saveFormToLedger')
    .addItem('พิมพ์ใบแจ้งอัตราค่าบริการ (PDF)', 'printRateStatementPdf')
    .addItem('พิมพ์ใบแจ้งค่าบริการ (PDF)', 'printServiceInvoicePdf')
    .addItem('พิมพ์เอกสารทั้ง 2 ใบ', 'printAllDocuments')
    .addSeparator()
    .addItem('โหลดรายการที่เลือกกลับมาพิมพ์', 'loadSelectedLedgerRecord')
    .addItem('ยกเลิกใบแจ้งหนี้แถวที่เลือก', 'cancelSelectedInvoice')
    .addItem('ซิงก์สถานะการชำระเงิน', 'syncPaymentStatuses')
    .addItem('รีเฟรชทะเบียนลูกหนี้', 'refreshLedgerSheetView')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function doGet(e) {
  const mode = String((e && e.parameter && e.parameter.mode) || '').trim().toLowerCase();
  if (mode === 'api') {
    return handleApiGet_(e);
  }

  const page = String((e && e.parameter && e.parameter.page) || 'dashboard').trim().toLowerCase();
  let fileName = 'Dashboard';
  let title = 'Dashboard ลูกหนี้';

  if (page === 'form') {
    fileName = 'InvoiceForm';
    title = 'เพิ่มใบแจ้งหนี้';
  } else if (page === 'detail') {
    fileName = 'DocumentDetail';
    title = 'รายละเอียดเอกสาร';
  }

  return HtmlService.createHtmlOutputFromFile(fileName)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const mode = String((e && e.parameter && e.parameter.mode) || '').trim().toLowerCase();
  if (mode !== 'api') {
    return createApiErrorOutput_(new Error('รองรับเฉพาะการเรียก POST แบบ API เท่านั้น'), e);
  }
  return handleApiPost_(e);
}

function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== APP_CONFIG.sheets.ledger || e.range.getRow() === 1) {
    return;
  }

  if (e.range.getColumn() !== LEDGER_COLUMNS.paidDate) {
    return;
  }

  try {
    updateSingleLedgerRowStatus_(sheet, e.range.getRow());
    refreshAllViews_();
  } catch (error) {
    console.error(error);
  }
}

function setupSystem() {
  ensureSystemReady_({ refreshViews: true });

  const bufferSheet = getSheet_(APP_CONFIG.sheets.form);
  if (!String(bufferSheet.getRange(APP_CONFIG.form.invoiceNumberCell).getDisplayValue()).trim()) {
    prepareNewInvoice();
  }

  safeToast_('ตั้งค่าระบบเรียบร้อยแล้ว', 'ระบบลูกหนี้', 5);
}

function prepareNewInvoice() {
  ensureSystemReady_({ refreshViews: false });
  const draft = createBlankDraft_();
  writeDraftToBuffer_(draft);
  safeToast_(`เตรียมเลขที่ใหม่ ${draft.invoiceNumber} แล้ว`, 'ระบบลูกหนี้', 5);
  return buildDraftResponse_(draft);
}

function saveFormToLedger() {
  const draft = readDraftFromBuffer_();
  const result = saveInvoiceData_(draft, { refreshViews: true });
  safeToast_(`บันทึกเลขที่ ${result.invoiceNumber} เรียบร้อย`, 'ระบบลูกหนี้', 5);
  return result;
}

function printRateStatementPdf() {
  const draft = readDraftFromBuffer_();
  const result = saveAndPrintFromDraft_(draft, 'rateStatement');
  SpreadsheetApp.getUi().alert(`สร้าง PDF ใบแจ้งอัตราค่าบริการเรียบร้อย\n${result.ratePdfUrl}`);
}

function printServiceInvoicePdf() {
  const draft = readDraftFromBuffer_();
  const result = saveAndPrintFromDraft_(draft, 'serviceInvoice');
  SpreadsheetApp.getUi().alert(`สร้าง PDF ใบแจ้งค่าบริการเรียบร้อย\n${result.invoicePdfUrl}`);
}

function printAllDocuments() {
  const draft = readDraftFromBuffer_();
  const result = saveAndPrintFromDraft_(draft, 'all');
  SpreadsheetApp.getUi().alert(
    `สร้าง PDF เรียบร้อยแล้ว\n\nใบแจ้งอัตราค่าบริการ:\n${result.ratePdfUrl}\n\nใบแจ้งค่าบริการ:\n${result.invoicePdfUrl}`
  );
}

function loadSelectedLedgerRecord() {
  const recordId = getRecordIdFromSelection_();
  const result = loadLedgerRecordById(recordId);
  safeToast_(`โหลดเลขที่ ${result.invoiceNumber} กลับเข้าฟอร์มระบบแล้ว`, 'ระบบลูกหนี้', 5);
  return result;
}

function loadLedgerRecordById(recordId) {
  ensureSystemReady_({ refreshViews: false });
  if (!recordId) {
    throw new Error('ไม่พบ Record ID สำหรับโหลดข้อมูล');
  }

  const row = findLedgerRowByRecordId_(recordId);
  if (!row) {
    throw new Error(`ไม่พบข้อมูล Record ID: ${recordId}`);
  }

  const draft = buildDraftFromLedgerRow_(row);
  writeDraftToBuffer_(draft);
  return buildDraftResponse_(draft, { row });
}

function cancelSelectedInvoice() {
  const recordId = getRecordIdFromSelection_();
  const row = findLedgerRowByRecordId_(recordId);
  if (!row) {
    throw new Error('ไม่พบแถวที่ต้องการยกเลิก');
  }

  const ledgerSheet = getSheet_(APP_CONFIG.sheets.ledger);
  const currentSystemStatus = String(ledgerSheet.getRange(row, LEDGER_COLUMNS.systemStatus).getValue()).trim();
  if (currentSystemStatus === APP_CONFIG.ledger.systemStatus.cancelled) {
    SpreadsheetApp.getUi().alert('รายการนี้ถูกยกเลิกไปแล้ว');
    return;
  }

  const response = SpreadsheetApp.getUi().prompt(
    'ยกเลิกใบแจ้งหนี้',
    'ระบุเหตุผลการยกเลิก (เว้นว่างได้)',
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
    return;
  }

  applyInvoiceCancellation_(row, response.getResponseText().trim());
  safeToast_('ยกเลิกใบแจ้งหนี้เรียบร้อย', 'ระบบลูกหนี้', 5);
}

function syncPaymentStatuses() {
  const ledgerSheet = getSheet_(APP_CONFIG.sheets.ledger);
  const lastRow = ledgerSheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  for (let row = 2; row <= lastRow; row += 1) {
    updateSingleLedgerRowStatus_(ledgerSheet, row);
  }

  refreshAllViews_();
  safeToast_('อัปเดตสถานะชำระเงินเรียบร้อย', 'ระบบลูกหนี้', 5);
}

function refreshDashboard() {
  refreshLedgerSheetView();
}

function refreshLedgerSheetView() {
  ensureSystemReady_({ refreshViews: true });
  safeToast_('รีเฟรชทะเบียนลูกหนี้เรียบร้อย', 'ระบบลูกหนี้', 5);
}

function showDashboard() {
  ensureSystemReady_({ refreshViews: false });
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Dashboard ลูกหนี้')
    .setWidth(480);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getWebAppBootstrap() {
  ensureSystemReady_({ refreshViews: false });
  return {
    maxRows: APP_CONFIG.form.maxRows,
    catalog: getCatalogOptions_(),
    draft: buildDraftResponse_(createBlankDraft_()),
    dashboard: getDashboardData(''),
    links: buildAppLinks_()
  };
}

function handleApiGet_(e) {
  try {
    ensureSystemReady_({ refreshViews: false });
    const action = getApiAction_(e);
    let data;

    switch (action) {
      case 'health':
        data = {
          status: 'ok',
          service: 'debtor-ledger-apps-script',
          timestamp: formatDateTimeForUi_(new Date()),
          maxRows: APP_CONFIG.form.maxRows,
          links: buildAppLinks_()
        };
        break;
      case 'bootstrap':
      case 'webbootstrap':
        data = getWebAppBootstrap();
        break;
      case 'dashboard':
      case 'list':
        data = getDashboardData(String((e && e.parameter && e.parameter.query) || ''));
        break;
      case 'form':
      case 'formbootstrap':
        data = getInvoiceFormBootstrap(String((e && e.parameter && e.parameter.recordId) || ''));
        break;
      case 'detail':
        data = getDocumentDetailBootstrap(String((e && e.parameter && e.parameter.recordId) || ''));
        break;
      case 'document':
      case 'documenturl':
        data = {
          url: getStoredDocumentUrl(
            String((e && e.parameter && e.parameter.recordId) || ''),
            String((e && e.parameter && e.parameter.type) || '')
          )
        };
        break;
      case 'new':
      case 'newinvoice':
        data = createNewInvoiceForWeb();
        break;
      default:
        throw new Error(`ไม่รู้จัก action ของ API: ${action}`);
    }

    return createApiSuccessOutput_(data, e);
  } catch (error) {
    console.error(error);
    return createApiErrorOutput_(error, e);
  }
}

function handleApiPost_(e) {
  try {
    ensureSystemReady_({ refreshViews: false });
    const action = getApiAction_(e);
    const body = parseApiBody_(e);
    const responseMode = getApiResponseMode_(e, body);
    let data;

    switch (action) {
      case 'save': {
        const payload = body && body.payload ? body.payload : body;
        data = saveInvoiceAndOpenDetailFromWeb(payload);
        break;
      }
      case 'cancel':
        data = cancelInvoiceFromWeb(
          String((body && body.recordId) || ''),
          String((body && body.reason) || '')
        );
        break;
      default:
        throw new Error(`ไม่รองรับ POST action นี้: ${action}`);
    }

    if (responseMode === 'redirect') {
      return createApiRedirectOutput_(action, data, e, body);
    }

    return createApiSuccessOutput_(data, e);
  } catch (error) {
    console.error(error);

    const body = parseApiBody_(e);
    const responseMode = getApiResponseMode_(e, body);
    if (responseMode === 'redirect') {
      return createApiRedirectErrorOutput_(error, e, body);
    }

    return createApiErrorOutput_(error, e);
  }
}

function getInvoiceFormBootstrap(recordId) {
  ensureSystemReady_({ refreshViews: false });

  let draft;
  if (recordId) {
    const row = findLedgerRowByRecordId_(recordId);
    if (!row) {
      throw new Error('ไม่พบข้อมูลที่ต้องการเปิดในฟอร์ม');
    }
    const detail = buildRecordDetailResponse_(row);
    draft = buildDraftResponse_(
      Object.assign({}, buildDraftFromLedgerRow_(row), {
        displayStatus: detail.displayStatus,
        systemStatus: detail.systemStatus,
        cancelReason: detail.cancelReason,
        cancelledAt: detail.cancelledAt,
        createdAt: detail.createdAt,
        updatedAt: detail.updatedAt,
        links: detail.links
      }),
      { row }
    );
  } else {
    draft = buildDraftResponse_(createBlankDraft_());
  }

  return {
    maxRows: APP_CONFIG.form.maxRows,
    catalog: getCatalogOptions_(),
    draft,
    links: buildAppLinks_()
  };
}

function getDocumentDetailBootstrap(recordId) {
  ensureSystemReady_({ refreshViews: false });
  if (!recordId) {
    throw new Error('ไม่พบรหัสรายการเอกสาร');
  }

  const row = findLedgerRowByRecordId_(recordId);
  if (!row) {
    throw new Error('ไม่พบข้อมูลเอกสารที่ต้องการ');
  }

  return {
    record: buildRecordDetailResponse_(row),
    links: buildAppLinks_()
  };
}

function saveInvoiceFromWeb(payload) {
  const result = saveInvoiceData_(payload, { refreshViews: true });
  return Object.assign({}, result, {
    draft: buildDraftResponse_(buildDraftFromLedgerRow_(result.row), { row: result.row }),
    dashboard: getDashboardData('')
  });
}

function saveInvoiceAndOpenDetailFromWeb(payload) {
  const result = saveAndPrintFromDraft_(payload, 'all');
  return {
    recordId: result.recordId,
    invoiceNumber: result.invoiceNumber,
    detailUrl: buildRecordDetailUrl_(result.recordId),
    record: buildRecordDetailResponse_(result.row)
  };
}

function saveAndPrintAllFromWeb(payload) {
  const result = saveAndPrintFromDraft_(payload, 'all');
  return Object.assign({}, result, {
    draft: buildDraftResponse_(buildDraftFromLedgerRow_(result.row), { row: result.row }),
    dashboard: getDashboardData('')
  });
}

function loadInvoiceIntoWebForm(recordId) {
  const draft = loadLedgerRecordById(recordId);
  return {
    draft,
    dashboard: getDashboardData('')
  };
}

function createNewInvoiceForWeb() {
  const draft = prepareNewInvoice();
  return {
    draft,
    dashboard: getDashboardData('')
  };
}

function cancelInvoiceFromWeb(recordId, reason) {
  const row = findLedgerRowByRecordId_(recordId);
  if (!row) {
    throw new Error('ไม่พบรายการที่ต้องการยกเลิก');
  }
  applyInvoiceCancellation_(row, String(reason || '').trim());
  return {
    dashboard: getDashboardData(''),
    draft: buildDraftResponse_(buildDraftFromLedgerRow_(row), { row })
  };
}

function getStoredDocumentUrl(recordId, type) {
  const row = findLedgerRowByRecordId_(recordId);
  if (!row) {
    throw new Error('ไม่พบรายการเอกสารที่ต้องการพิมพ์');
  }

  const record = getLedgerRecords_().find((item) => item.row === row);
  if (!record) {
    throw new Error('ไม่พบข้อมูลสำหรับสร้างเอกสาร');
  }

  if (type === 'rateStatement' && record.ratePdfUrl) {
    return record.ratePdfUrl;
  }
  if (type === 'serviceInvoice' && record.invoicePdfUrl) {
    return record.invoicePdfUrl;
  }

  const context = {
    row,
    recordId: record.recordId,
    invoiceNumber: record.invoiceNumber,
    customerName: record.customerName,
    ratePdfUrl: record.ratePdfUrl,
    invoicePdfUrl: record.invoicePdfUrl
  };
  const exported = exportDocumentForLedgerRow_(context, type);
  const links = type === 'rateStatement' ? { ratePdfUrl: exported.url } : { invoicePdfUrl: exported.url };
  updateLedgerDocumentLinks_(row, links);
  return exported.url;
}

function getDashboardData(query) {
  const records = getLedgerRecords_(query);
  const activeRecords = records.filter((record) => !record.isCancelled);
  const totalAmount = sumBy_(activeRecords, 'amount');
  const paidAmount = sumBy_(records, 'paidAmount');
  const outstandingAmount = sumBy_(records, 'outstandingAmount');
  const paidCount = records.filter((record) => record.isPaid && !record.isCancelled).length;
  const pendingCount = records.filter((record) => !record.isPaid && !record.isCancelled).length;
  const cancelledCount = records.filter((record) => record.isCancelled).length;
  const collectionRate = totalAmount > 0 ? (paidAmount / totalAmount) * 100 : 0;
  const documentPaidRate = activeRecords.length > 0 ? (paidCount / activeRecords.length) * 100 : 0;
  const outstandingRate = totalAmount > 0 ? (outstandingAmount / totalAmount) * 100 : 0;
  const monthlyTrend = buildDashboardMonthlyTrend_(records, 8);
  const statusBreakdown = buildDashboardStatusBreakdown_(records, {
    totalAmount,
    paidAmount,
    outstandingAmount,
    pendingCount,
    paidCount,
    cancelledCount
  });
  const topDebtors = buildDashboardTopDebtors_(records, 6);
  const latestRecord = buildDashboardLatestRecord_(records);
  const monthlyDocStats = monthlyTrend.slice(-6);

  return {
    summary: {
      totalAmount,
      paidAmount,
      outstandingAmount,
      invoiceCount: activeRecords.length,
      paidCount,
      pendingCount,
      cancelledCount,
      collectionRate,
      documentPaidRate,
      outstandingRate,
      averageAmount: activeRecords.length > 0 ? totalAmount / activeRecords.length : 0
    },
    monthlyTrend,
    monthlyDocStats,
    statusBreakdown,
    topDebtors,
    latestRecord,
    ringMetrics: [
      { label: 'อัตราการเก็บหนี้', value: collectionRate, tone: 'debt' },
      { label: 'เอกสารชำระแล้ว', value: documentPaidRate, tone: 'paid' },
      { label: 'สัดส่วนยอดค้าง', value: outstandingRate, tone: 'outstanding' }
    ],
    outstanding: records
      .filter((record) => record.outstandingAmount > 0)
      .sort((a, b) => dateValue_(b.invoiceDate) - dateValue_(a.invoiceDate))
      .slice(0, 12)
      .map(mapRecordForUi_),
    recent: records
      .sort((a, b) => dateValue_(b.updatedAt || b.invoiceDate) - dateValue_(a.updatedAt || a.invoiceDate))
      .slice(0, 12)
      .map(mapRecordForUi_),
    lastUpdated: formatDateTimeForUi_(new Date()),
    query: String(query || '').trim(),
    links: buildAppLinks_()
  };
}

function buildDashboardMonthlyTrend_(records, monthCount) {
  const months = [];
  const bucketMap = {};
  const now = new Date();

  for (let index = monthCount - 1; index >= 0; index -= 1) {
    const monthDate = new Date(now.getFullYear(), now.getMonth() - index, 1);
    const key = Utilities.formatDate(monthDate, APP_CONFIG.timezone, 'yyyy-MM');
    const bucket = {
      key,
      label: formatMonthLabelShort_(monthDate),
      issuedAmount: 0,
      paidAmount: 0,
      outstandingAmount: 0,
      documentCount: 0,
      paidDocumentCount: 0
    };
    months.push(bucket);
    bucketMap[key] = bucket;
  }

  records.forEach((record) => {
    const invoiceDate = toDateOrNull_(record.invoiceDate);
    if (invoiceDate) {
      const key = Utilities.formatDate(invoiceDate, APP_CONFIG.timezone, 'yyyy-MM');
      if (bucketMap[key]) {
        bucketMap[key].issuedAmount += toNumber_(record.amount);
        bucketMap[key].outstandingAmount += toNumber_(record.outstandingAmount);
        bucketMap[key].documentCount += 1;
      }
    }

    const paidDate = toDateOrNull_(record.paidDate);
    if (record.isPaid && paidDate) {
      const key = Utilities.formatDate(paidDate, APP_CONFIG.timezone, 'yyyy-MM');
      if (bucketMap[key]) {
        bucketMap[key].paidAmount += toNumber_(record.amount);
        bucketMap[key].paidDocumentCount += 1;
      }
    }
  });

  return months.map((bucket) =>
    Object.assign({}, bucket, {
      issuedAmount: roundToTwo_(bucket.issuedAmount),
      paidAmount: roundToTwo_(bucket.paidAmount),
      outstandingAmount: roundToTwo_(bucket.outstandingAmount)
    })
  );
}

function buildDashboardStatusBreakdown_(records, totals) {
  const cancelledAmount = records
    .filter((record) => record.isCancelled)
    .reduce((sum, record) => sum + toNumber_(record.amount), 0);

  return {
    totalCount: totals.pendingCount + totals.paidCount + totals.cancelledCount,
    totalAmount: roundToTwo_(totals.totalAmount),
    items: [
      {
        key: 'pending',
        label: 'รอชำระ',
        count: totals.pendingCount,
        amount: roundToTwo_(totals.outstandingAmount),
        color: '#f97316'
      },
      {
        key: 'paid',
        label: 'ชำระแล้ว',
        count: totals.paidCount,
        amount: roundToTwo_(totals.paidAmount),
        color: '#22c55e'
      },
      {
        key: 'cancelled',
        label: 'ยกเลิก',
        count: totals.cancelledCount,
        amount: roundToTwo_(cancelledAmount),
        color: '#ef4444'
      }
    ]
  };
}

function buildDashboardTopDebtors_(records, limit) {
  const grouped = {};

  records
    .filter((record) => record.outstandingAmount > 0 && !record.isCancelled)
    .forEach((record) => {
      const key = String(record.customerName || 'ไม่ระบุชื่อ').trim() || 'ไม่ระบุชื่อ';
      if (!grouped[key]) {
        grouped[key] = {
          customerName: key,
          outstandingAmount: 0,
          documentCount: 0,
          latestInvoiceDate: null,
          latestInvoiceNumber: '',
          latestRecordId: ''
        };
      }
      grouped[key].outstandingAmount += toNumber_(record.outstandingAmount);
      grouped[key].documentCount += 1;

      const invoiceDate = toDateOrNull_(record.invoiceDate);
      if (invoiceDate && (!grouped[key].latestInvoiceDate || invoiceDate > grouped[key].latestInvoiceDate)) {
        grouped[key].latestInvoiceDate = invoiceDate;
        grouped[key].latestInvoiceNumber = record.invoiceNumber || '';
        grouped[key].latestRecordId = record.recordId || '';
      }
    });

  const totalOutstanding = Object.values(grouped).reduce((sum, item) => sum + item.outstandingAmount, 0);

  return Object.values(grouped)
    .sort((a, b) => b.outstandingAmount - a.outstandingAmount)
    .slice(0, limit)
    .map((item) => ({
      customerName: item.customerName,
      outstandingAmount: roundToTwo_(item.outstandingAmount),
      documentCount: item.documentCount,
      latestInvoiceNumber: item.latestInvoiceNumber,
      latestRecordId: item.latestRecordId,
      latestInvoiceDate: formatDateForUi_(item.latestInvoiceDate),
      share: totalOutstanding > 0 ? roundToTwo_((item.outstandingAmount / totalOutstanding) * 100) : 0
    }));
}

function buildDashboardLatestRecord_(records) {
  if (!records.length) {
    return null;
  }
  const latestRecord = records
    .slice()
    .sort((a, b) => dateValue_(b.updatedAt || b.invoiceDate) - dateValue_(a.updatedAt || a.invoiceDate))[0];

  return latestRecord
    ? Object.assign({}, mapRecordForUi_(latestRecord), {
        updatedAt: formatDateTimeForUi_(latestRecord.updatedAt || latestRecord.invoiceDate)
      })
    : null;
}

function saveAndPrintFromDraft_(payload, mode) {
  const context = saveInvoiceData_(payload, { refreshViews: false });
  const links = {};

  if (mode === 'rateStatement' || mode === 'all') {
    const rate = exportDocumentForLedgerRow_(context, 'rateStatement');
    links.ratePdfUrl = rate.url;
  }

  if (mode === 'serviceInvoice' || mode === 'all') {
    const invoice = exportDocumentForLedgerRow_(context, 'serviceInvoice');
    links.invoicePdfUrl = invoice.url;
  }

  updateLedgerDocumentLinks_(context.row, links);
  refreshAllViews_();

  const latestRowDraft = buildDraftFromLedgerRow_(context.row);
  return Object.assign({}, context, links, {
    draft: buildDraftResponse_(latestRowDraft, { row: context.row })
  });
}

function saveInvoiceData_(payload, options) {
  const settings = Object.assign({ refreshViews: true }, options || {});
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    ensureSystemReady_({ refreshViews: false });

    const draft = normalizeDraftInput_(payload);
    validateFormData_(draft);

    const ledgerSheet = getSheet_(APP_CONFIG.sheets.ledger);
    const existingRowByRecordId = draft.recordId ? findLedgerRowByRecordId_(draft.recordId) : 0;

    let invoiceNumber = draft.invoiceNumber || generateNextInvoiceNumber_(draft.issueDate);
    let duplicateRowByNumber = findLedgerRowByInvoiceNumber_(invoiceNumber);

    if (!existingRowByRecordId && duplicateRowByNumber) {
      invoiceNumber = generateNextInvoiceNumber_(draft.issueDate);
      duplicateRowByNumber = findLedgerRowByInvoiceNumber_(invoiceNumber);
    }

    if (duplicateRowByNumber && duplicateRowByNumber !== existingRowByRecordId) {
      throw new Error(`เลขที่ใบแจ้ง ${invoiceNumber} ถูกใช้งานแล้ว กรุณาสร้างรายการใหม่อีกครั้ง`);
    }

    draft.invoiceNumber = invoiceNumber;
    writeDraftToBuffer_(draft);
    SpreadsheetApp.flush();

    const rateSheet = getSheet_(APP_CONFIG.sheets.rateStatement);
    const serviceInvoiceSheet = getSheet_(APP_CONFIG.sheets.serviceInvoice);
    const rateTotal = toNumber_(rateSheet.getRange(APP_CONFIG.totalCells.rateStatement).getValue());
    const invoiceTotal = toNumber_(serviceInvoiceSheet.getRange(APP_CONFIG.totalCells.serviceInvoice).getValue());
    const amount = invoiceTotal;

    if (amount <= 0) {
      throw new Error('ไม่พบยอดในใบแจ้งค่าบริการ กรุณากรอกรายการฝั่งใบแจ้งค่าบริการก่อนบันทึก เพราะระบบใช้ยอดนี้เป็นยอดลูกหนี้');
    }

    const row = existingRowByRecordId || ledgerSheet.getLastRow() + 1;
    const existingValues = existingRowByRecordId
      ? ledgerSheet.getRange(existingRowByRecordId, 1, 1, LEDGER_COLUMNS.lastColumn).getValues()[0]
      : null;

    if (
      existingValues &&
      String(existingValues[LEDGER_COLUMNS.systemStatus - 1]).trim() === APP_CONFIG.ledger.systemStatus.cancelled
    ) {
      throw new Error('รายการนี้ถูกยกเลิกแล้ว ไม่สามารถบันทึกทับได้');
    }

    if (existingValues) {
      const currentRevisionToken = buildRevisionToken_(
        toDateOrNull_(existingValues[LEDGER_COLUMNS.updatedAt - 1]) ||
        toDateOrNull_(existingValues[LEDGER_COLUMNS.createdAt - 1])
      );
      if (draft.revisionToken && currentRevisionToken && draft.revisionToken !== currentRevisionToken) {
        throw new Error('ข้อมูลรายการนี้ถูกแก้ไขจากเครื่องอื่นแล้ว กรุณาเปิดรายการใหม่อีกครั้งก่อนบันทึกเพื่อป้องกันข้อมูลทับกัน');
      }
    }

    const now = new Date();
    const recordId = draft.recordId || buildRecordId_(draft.invoiceNumber);
    const paidDate = existingValues ? existingValues[LEDGER_COLUMNS.paidDate - 1] : '';
    const systemStatus = existingValues
      ? String(existingValues[LEDGER_COLUMNS.systemStatus - 1]).trim() || APP_CONFIG.ledger.systemStatus.active
      : APP_CONFIG.ledger.systemStatus.active;
    const receiptStatus = computeDisplayStatus_(systemStatus, paidDate);
    const outstandingAmount = receiptStatus === APP_CONFIG.ledger.displayStatus.pending ? amount : 0;

    const rowValues = [
      draft.customerName,
      draft.customerAddress,
      draft.issueDate,
      draft.invoiceNumber,
      draft.issueDate,
      amount,
      paidDate,
      receiptStatus,
      recordId,
      systemStatus,
      JSON.stringify(trimEmptyItems_(draft.rateItems)),
      JSON.stringify(trimEmptyItems_(draft.invoiceItems)),
      existingValues ? existingValues[LEDGER_COLUMNS.ratePdfUrl - 1] : '',
      existingValues ? existingValues[LEDGER_COLUMNS.invoicePdfUrl - 1] : '',
      rateTotal,
      invoiceTotal,
      existingValues ? existingValues[LEDGER_COLUMNS.createdAt - 1] || now : now,
      now,
      existingValues ? existingValues[LEDGER_COLUMNS.cancelledAt - 1] : '',
      existingValues ? existingValues[LEDGER_COLUMNS.cancelReason - 1] : '',
      outstandingAmount
    ];

    ledgerSheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
    updateSingleLedgerRowStatus_(ledgerSheet, row);

    writeDraftToBuffer_(
      Object.assign({}, draft, {
        recordId,
        row
      })
    );

    if (settings.refreshViews) {
      refreshAllViews_();
    }

    return {
      row,
      recordId,
      invoiceNumber: draft.invoiceNumber,
      customerName: draft.customerName,
      amount,
      rateTotal,
      invoiceTotal,
      ratePdfUrl: existingValues ? existingValues[LEDGER_COLUMNS.ratePdfUrl - 1] : '',
      invoicePdfUrl: existingValues ? existingValues[LEDGER_COLUMNS.invoicePdfUrl - 1] : ''
    };
  } finally {
    lock.releaseLock();
  }
}

function refreshAllViews_() {
  prepareLedgerSheet_();
}

function ensureSystemReady_(options) {
  const settings = Object.assign({ refreshViews: false }, options || {});
  assertRequiredSheets_();
  prepareLedgerSheet_();
  prepareFormBufferSheet_();
  syncPrintTemplateFormulas_();
  if (settings.refreshViews) {
    refreshAllViews_();
  }
}

function prepareLedgerSheet_() {
  const sheet = getSheet_(APP_CONFIG.sheets.ledger);
  const headers = APP_CONFIG.ledger.coreHeaders.concat(APP_CONFIG.ledger.extraHeaders);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  const widths = {
    1: 220,
    2: 320,
    3: 130,
    4: 150,
    5: 150,
    6: 130,
    7: 130,
    8: 150,
    9: 220,
    10: 140,
    13: 190,
    14: 190,
    15: 140,
    16: 140,
    17: 140,
    18: 150,
    19: 150,
    20: 220,
    21: 140
  };
  Object.keys(widths).forEach((column) => sheet.setColumnWidth(Number(column), widths[column]));
  sheet.showColumns(LEDGER_COLUMNS.recordId, LEDGER_COLUMNS.lastColumn - LEDGER_COLUMNS.recordId + 1);
  sheet.hideColumns(LEDGER_COLUMNS.rateItemsJson, 2);

  sheet
    .getRange(1, 1, 1, headers.length)
    .setBackground(APP_CONFIG.colors.header)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, LEDGER_COLUMNS.invoiceDate, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(2, LEDGER_COLUMNS.rateDate, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(2, LEDGER_COLUMNS.paidDate, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(2, LEDGER_COLUMNS.amount, lastRow - 1, 1).setNumberFormat('#,##0.00');
    sheet.getRange(2, LEDGER_COLUMNS.rateTotal, lastRow - 1, 2).setNumberFormat('#,##0.00');
    sheet.getRange(2, LEDGER_COLUMNS.outstandingAmount, lastRow - 1, 1).setNumberFormat('#,##0.00');
    sheet.getRange(2, LEDGER_COLUMNS.createdAt, lastRow - 1, 3).setNumberFormat('dd/MM/yyyy HH:mm');
    for (let row = 2; row <= lastRow; row += 1) {
      updateSingleLedgerRowStatus_(sheet, row);
    }
  }
}

function prepareFormBufferSheet_() {
  const sheet = getOrCreateSheet_(APP_CONFIG.sheets.form);

  sheet.getRange('A1').setValue('ชีตระบบสำหรับเก็บข้อมูลฟอร์มชั่วคราว ห้ามกรอกโดยตรง');
  sheet.getRange('A2').setValue('ระบบจะใช้ชีตนี้เป็นตัวกลางสำหรับพิมพ์แบบฟอร์ม');
  sheet.getRange('B3:C22').setNumberFormat('@');
  sheet.getRange('F3:G22').setNumberFormat('@');
  sheet.getRange('D3:D22').setNumberFormat('#,##0.00');
  sheet.getRange('H3:H22').setNumberFormat('#,##0.00');
  sheet.getRange('K3:K7').setNumberFormat('@');

  if (!sheet.isSheetHidden()) {
    const activeSheet = getSpreadsheet_().getActiveSheet();
    if (activeSheet && activeSheet.getSheetId() === sheet.getSheetId()) {
      getSheet_(APP_CONFIG.sheets.ledger).activate();
    }
    sheet.hideSheet();
  }
}

function syncPrintTemplateFormulas_() {
  const bufferSheetName = APP_CONFIG.sheets.form;
  const rateSheet = getSheet_(APP_CONFIG.sheets.rateStatement);
  const invoiceSheet = getSheet_(APP_CONFIG.sheets.serviceInvoice);

  rateSheet.getRange('B10').setFormula(`="เรียน/To : "&${sheetRef_(bufferSheetName, 'K6')}`);
  rateSheet.getRange('C11').setFormula(`=${sheetRef_(bufferSheetName, 'K7')}`);
  rateSheet.getRange('G12').setFormula(`=${sheetRef_(bufferSheetName, 'K3')}`);
  rateSheet.getRange('G14').setFormula(`="วันที่ :            "&${sheetRef_(bufferSheetName, 'K4')}`);
  rateSheet.getRange('G15').setFormula(`="Date :           "&${sheetRef_(bufferSheetName, 'K5')}`);
  rateSheet.getRange('G14:G15').setNumberFormat('@');

  invoiceSheet.getRange('B10').setFormula(`="เรียน/To : "&${sheetRef_(bufferSheetName, 'K6')}`);
  invoiceSheet.getRange('C11').setFormula(`=${sheetRef_(bufferSheetName, 'K7')}`);
  invoiceSheet.getRange('G12').setFormula(`='${APP_CONFIG.sheets.rateStatement}'!G12`);
  invoiceSheet.getRange('G14').setFormula(`='${APP_CONFIG.sheets.rateStatement}'!G14`);
  invoiceSheet.getRange('G15').setFormula(`='${APP_CONFIG.sheets.rateStatement}'!G15`);
  invoiceSheet.getRange('G14:G15').setNumberFormat('@');

  for (let index = 0; index < APP_CONFIG.form.maxRows; index += 1) {
    const formRow = index + 3;
    const printRow = index + 18;

    rateSheet
      .getRange(`B${printRow}`)
      .setFormula(`=IF(C${printRow}<>"",COUNTA($C$18:C${printRow}),"")`);
    rateSheet
      .getRange(`C${printRow}`)
      .setFormula(`=IF(ISBLANK(${sheetRef_(bufferSheetName, `B${formRow}`)}),"",${sheetRef_(bufferSheetName, `B${formRow}`)})`);
    rateSheet
      .getRange(`D${printRow}`)
      .setFormula(`=IF(ISBLANK(C${printRow}),"",XLOOKUP(C${printRow},'${APP_CONFIG.sheets.catalog}'!A:A,'${APP_CONFIG.sheets.catalog}'!B:B,"",0))`);
    rateSheet
      .getRange(`E${printRow}`)
      .setFormula(`=IF(ISBLANK(${sheetRef_(bufferSheetName, `C${formRow}`)}),"",${sheetRef_(bufferSheetName, `C${formRow}`)})`);
    rateSheet
      .getRange(`F${printRow}`)
      .setFormula(`=IF(ISBLANK(${sheetRef_(bufferSheetName, `D${formRow}`)}),"",${sheetRef_(bufferSheetName, `D${formRow}`)})`);
    rateSheet
      .getRange(`G${printRow}`)
      .setFormula(`=IF(C${printRow}="","",XLOOKUP(C${printRow},'${APP_CONFIG.sheets.catalog}'!A:A,'${APP_CONFIG.sheets.catalog}'!E:E,0,0))`);
    rateSheet
      .getRange(`H${printRow}`)
      .setFormula(`=IF(C${printRow}="","",F${printRow}*G${printRow})`);

    invoiceSheet
      .getRange(`B${printRow}`)
      .setFormula(`=IF(C${printRow}<>"",COUNTA($C$18:C${printRow}),"")`);
    invoiceSheet
      .getRange(`C${printRow}`)
      .setFormula(`=IF(ISBLANK(${sheetRef_(bufferSheetName, `F${formRow}`)}),"",${sheetRef_(bufferSheetName, `F${formRow}`)})`);
    invoiceSheet
      .getRange(`D${printRow}`)
      .setFormula(`=IF(ISBLANK(C${printRow}),"",XLOOKUP(C${printRow},'${APP_CONFIG.sheets.catalog}'!A:A,'${APP_CONFIG.sheets.catalog}'!B:B,"",0))`);
    invoiceSheet
      .getRange(`E${printRow}`)
      .setFormula(`=IF(ISBLANK(${sheetRef_(bufferSheetName, `G${formRow}`)}),"",${sheetRef_(bufferSheetName, `G${formRow}`)})`);
    invoiceSheet
      .getRange(`F${printRow}`)
      .setFormula(`=IF(ISBLANK(${sheetRef_(bufferSheetName, `H${formRow}`)}),"",${sheetRef_(bufferSheetName, `H${formRow}`)})`);
    invoiceSheet
      .getRange(`G${printRow}`)
      .setFormula(`=IF(C${printRow}="","",XLOOKUP(C${printRow},'${APP_CONFIG.sheets.catalog}'!A:A,'${APP_CONFIG.sheets.catalog}'!E:E,0,0))`);
    invoiceSheet
      .getRange(`H${printRow}`)
      .setFormula(`=IF(C${printRow}="","",F${printRow}*G${printRow})`);
  }

  rateSheet.getRange('H38').setFormula('=SUM(H18:H37)');
  rateSheet.getRange('B39').setFormula('="("&BAHTTEXT(H38)&")"');
  invoiceSheet.getRange('H38').setFormula('=SUM(H18:H37)');
  invoiceSheet.getRange('B39').setFormula('="("&BAHTTEXT(H38)&")"');
}

function refreshDashboardSheet_(records) {
  const sheet = getOrCreateSheet_(APP_CONFIG.sheets.dashboard);
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
  sheet.getDataRange().breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(0);

  for (let column = 1; column <= 12; column += 1) {
    sheet.setColumnWidth(column, column === 1 ? 180 : 120);
  }

  sheet.getRange('A1:L2').merge();
  sheet
    .getRange('A1')
    .setValue('Dashboard ระบบทะเบียนคุมลูกหนี้')
    .setBackground(APP_CONFIG.colors.header)
    .setFontColor('#ffffff')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const activeRecords = records.filter((record) => !record.isCancelled);
  const totalAmount = sumBy_(activeRecords, 'amount');
  const paidAmount = sumBy_(records, 'paidAmount');
  const outstandingAmount = sumBy_(records, 'outstandingAmount');
  const activeCount = activeRecords.length;
  const pendingCount = records.filter((record) => !record.isPaid && !record.isCancelled).length;
  const paidCount = records.filter((record) => record.isPaid && !record.isCancelled).length;
  const cancelledCount = records.filter((record) => record.isCancelled).length;

  [
    { range: 'A4:C6', title: 'ยอดลูกหนี้รวม', value: totalAmount, color: '#dbeafe', accent: '#1d4ed8' },
    { range: 'D4:F6', title: 'ยอดชำระแล้ว', value: paidAmount, color: '#dcfce7', accent: '#15803d' },
    { range: 'G4:I6', title: 'ยอดค้างชำระ', value: outstandingAmount, color: '#ffedd5', accent: '#ea580c' },
    { range: 'J4:L6', title: 'จำนวนรายการใช้งาน', value: String(activeCount), color: '#ede9fe', accent: '#6d28d9' }
  ].forEach((card) => {
    sheet.getRange(card.range).merge();
    sheet
      .getRange(card.range.split(':')[0])
      .setValue(`${card.title}\n${formatNumberForSheet_(card.value)}`)
      .setBackground(card.color)
      .setFontColor(card.accent)
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);
  });

  sheet.getRange('A8:F8').merge();
  sheet.getRange('G8:L8').merge();
  sheet
    .getRange('A8')
    .setValue(`สถานะรวม: รอชำระ ${pendingCount} | ชำระแล้ว ${paidCount} | ยกเลิก ${cancelledCount}`)
    .setBackground(APP_CONFIG.colors.brandSoft)
    .setFontWeight('bold');
  sheet
    .getRange('G8')
    .setValue(
      `อัตราการเก็บหนี้: ${totalAmount > 0 ? ((paidAmount / totalAmount) * 100).toFixed(2) : '0.00'}%`
    )
    .setBackground(APP_CONFIG.colors.brandSoft)
    .setFontWeight('bold');

  const outstandingHeader = [['ลูกหนี้ค้างชำระ', 'เลขที่', 'วันที่ออก', 'จำนวนเงิน', 'สถานะ', 'PDF']];
  sheet.getRange(10, 1, 1, outstandingHeader[0].length).setValues(outstandingHeader);
  styleTableHeader_(sheet.getRange(10, 1, 1, outstandingHeader[0].length));

  const outstandingRows = records
    .filter((record) => record.outstandingAmount > 0)
    .sort((a, b) => dateValue_(b.invoiceDate) - dateValue_(a.invoiceDate))
    .slice(0, 15)
    .map((record) => [
      record.customerName,
      record.invoiceNumber,
      record.invoiceDate || '',
      record.outstandingAmount,
      record.displayStatus,
      record.invoicePdfUrl ? `=HYPERLINK("${record.invoicePdfUrl}","เปิด PDF")` : ''
    ]);

  if (outstandingRows.length) {
    sheet.getRange(11, 1, outstandingRows.length, outstandingRows[0].length).setValues(outstandingRows);
    sheet.getRange(11, 3, outstandingRows.length, 1).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(11, 4, outstandingRows.length, 1).setNumberFormat('#,##0.00');
  } else {
    sheet.getRange('A11:F11').merge().setValue('ยังไม่มีรายการค้างชำระ');
  }

  const recentStartColumn = 8;
  const recentHeader = [['รายการล่าสุด', 'เลขที่', 'ยอดเงิน', 'ชำระ', 'สถานะ']];
  sheet.getRange(10, recentStartColumn, 1, recentHeader[0].length).setValues(recentHeader);
  styleTableHeader_(sheet.getRange(10, recentStartColumn, 1, recentHeader[0].length));

  const recentRows = records
    .sort((a, b) => dateValue_(b.updatedAt || b.invoiceDate) - dateValue_(a.updatedAt || a.invoiceDate))
    .slice(0, 15)
    .map((record) => [record.customerName, record.invoiceNumber, record.amount, record.paidDate || '', record.displayStatus]);

  if (recentRows.length) {
    sheet.getRange(11, recentStartColumn, recentRows.length, recentRows[0].length).setValues(recentRows);
    sheet.getRange(11, recentStartColumn + 2, recentRows.length, 1).setNumberFormat('#,##0.00');
    sheet.getRange(11, recentStartColumn + 3, recentRows.length, 1).setNumberFormat('dd/MM/yyyy');
  } else {
    sheet.getRange(11, recentStartColumn, 1, 5).merge().setValue('ยังไม่มีข้อมูลล่าสุด');
  }
}

function refreshReceiptSummarySheet_(records) {
  const sheet = getOrCreateSheet_(APP_CONFIG.sheets.receiptSummary);
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
  sheet.getDataRange().breakApart();
  sheet.clear();
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(false);

  const headers = [
    'Record ID',
    'เลขที่ใบแจ้ง',
    'ชื่อผู้ใช้บริการ',
    'วันที่ออกใบแจ้ง',
    'จำนวนเงิน',
    'วันที่ชำระ',
    'สถานะ',
    'PDF ใบแจ้งอัตราค่าบริการ',
    'PDF ใบแจ้งค่าบริการ',
    'เหตุผลยกเลิก'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleTableHeader_(sheet.getRange(1, 1, 1, headers.length));

  if (!records.length) {
    sheet.getRange('A2:J2').merge().setValue('ยังไม่มีข้อมูลในระบบ');
    return;
  }

  const rows = records
    .sort((a, b) => dateValue_(b.invoiceDate) - dateValue_(a.invoiceDate))
    .map((record) => [
      record.recordId,
      record.invoiceNumber,
      record.customerName,
      record.invoiceDate || '',
      record.amount,
      record.paidDate || '',
      record.displayStatus,
      record.ratePdfUrl ? `=HYPERLINK("${record.ratePdfUrl}","เปิดไฟล์")` : '',
      record.invoicePdfUrl ? `=HYPERLINK("${record.invoicePdfUrl}","เปิดไฟล์")` : '',
      record.cancelReason || ''
    ]);

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(2, 4, rows.length, 1).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 6, rows.length, 1).setNumberFormat('dd/MM/yyyy');

  [180, 130, 220, 120, 120, 120, 140, 150, 130, 220].forEach((width, index) =>
    sheet.setColumnWidth(index + 1, width)
  );

  if (sheet.getLastRow() > 1) {
    sheet.getRange(1, 1, sheet.getLastRow(), headers.length).createFilter();
  }
}

function getLedgerRecords_(query) {
  const ledgerSheet = getSheet_(APP_CONFIG.sheets.ledger);
  const lastRow = ledgerSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const values = ledgerSheet.getRange(2, 1, lastRow - 1, LEDGER_COLUMNS.lastColumn).getValues();
  const search = String(query || '').trim().toLowerCase();

  return values
    .map((row, index) => {
      const amount = toNumber_(row[LEDGER_COLUMNS.amount - 1]);
      const systemStatus = String(row[LEDGER_COLUMNS.systemStatus - 1]).trim() || APP_CONFIG.ledger.systemStatus.active;
      const paidDate = toDateOrNull_(row[LEDGER_COLUMNS.paidDate - 1]);
      const isCancelled = systemStatus === APP_CONFIG.ledger.systemStatus.cancelled;
      const isPaid = Boolean(paidDate) && !isCancelled;
      return {
        row: index + 2,
        customerName: row[LEDGER_COLUMNS.customerName - 1],
        address: row[LEDGER_COLUMNS.address - 1],
        invoiceDate: toDateOrNull_(row[LEDGER_COLUMNS.invoiceDate - 1]),
        invoiceNumber: row[LEDGER_COLUMNS.rateNumber - 1],
        rateDate: toDateOrNull_(row[LEDGER_COLUMNS.rateDate - 1]),
        amount,
        paidDate,
        displayStatus: computeDisplayStatus_(systemStatus, paidDate),
        recordId: row[LEDGER_COLUMNS.recordId - 1],
        systemStatus,
        ratePdfUrl: row[LEDGER_COLUMNS.ratePdfUrl - 1],
        invoicePdfUrl: row[LEDGER_COLUMNS.invoicePdfUrl - 1],
        rateTotal: toNumber_(row[LEDGER_COLUMNS.rateTotal - 1]),
        invoiceTotal: toNumber_(row[LEDGER_COLUMNS.invoiceTotal - 1]),
        createdAt: toDateOrNull_(row[LEDGER_COLUMNS.createdAt - 1]),
        updatedAt: toDateOrNull_(row[LEDGER_COLUMNS.updatedAt - 1]),
        cancelledAt: toDateOrNull_(row[LEDGER_COLUMNS.cancelledAt - 1]),
        cancelReason: row[LEDGER_COLUMNS.cancelReason - 1],
        isCancelled,
        isPaid,
        paidAmount: isPaid ? amount : 0,
        outstandingAmount: isCancelled || isPaid ? 0 : amount
      };
    })
    .filter((record) => {
      if (!search) {
        return true;
      }
      const haystack = [record.customerName, record.invoiceNumber, record.recordId, record.displayStatus].join(' ').toLowerCase();
      return haystack.includes(search);
    });
}

function assertRequiredSheets_() {
  [
    APP_CONFIG.sheets.rateStatement,
    APP_CONFIG.sheets.serviceInvoice,
    APP_CONFIG.sheets.catalog,
    APP_CONFIG.sheets.ledger
  ].forEach((sheetName) => {
    if (!getSpreadsheet_().getSheetByName(sheetName)) {
      throw new Error(`ไม่พบชีต "${sheetName}" กรุณาตรวจสอบชื่อชีตให้ตรงกับต้นฉบับ`);
    }
  });
}

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`ไม่พบชีต ${sheetName}`);
  }
  return sheet;
}

function getOrCreateSheet_(sheetName) {
  return getSpreadsheet_().getSheetByName(sheetName) || getSpreadsheet_().insertSheet(sheetName);
}

function createBlankDraft_() {
  const issueDate = new Date();
  return {
    recordId: '',
    row: '',
    revisionToken: '',
    invoiceNumber: generateNextInvoiceNumber_(issueDate),
    issueDate,
    customerName: '',
    customerAddress: '',
    rateItems: createEmptyItems_(),
    invoiceItems: createEmptyItems_()
  };
}

function buildDraftFromLedgerRow_(row) {
  const values = getSheet_(APP_CONFIG.sheets.ledger).getRange(row, 1, 1, LEDGER_COLUMNS.lastColumn).getValues()[0];
  const createdAt = toDateOrNull_(values[LEDGER_COLUMNS.createdAt - 1]);
  const updatedAt = toDateOrNull_(values[LEDGER_COLUMNS.updatedAt - 1]);
  return {
    recordId: values[LEDGER_COLUMNS.recordId - 1] || '',
    row,
    revisionToken: buildRevisionToken_(updatedAt || createdAt),
    invoiceNumber: values[LEDGER_COLUMNS.rateNumber - 1] || '',
    issueDate: toDateOrNull_(values[LEDGER_COLUMNS.invoiceDate - 1]) || new Date(),
    customerName: values[LEDGER_COLUMNS.customerName - 1] || '',
    customerAddress: values[LEDGER_COLUMNS.address - 1] || '',
    rateItems: padItems_(parseJsonArray_(values[LEDGER_COLUMNS.rateItemsJson - 1])),
    invoiceItems: padItems_(parseJsonArray_(values[LEDGER_COLUMNS.invoiceItemsJson - 1]))
  };
}

function buildDraftResponse_(draft, extra) {
  const normalized = normalizeDraftInput_(draft);
  const source = draft || {};
  return Object.assign(
    {
      recordId: normalized.recordId || '',
      row: normalized.row || '',
      revisionToken: normalized.revisionToken || '',
      invoiceNumber: normalized.invoiceNumber || '',
      issueDate: formatDateInputValue_(normalized.issueDate),
      issueDateThai: formatThaiDate_(normalized.issueDate),
      issueDateEnglish: formatEnglishDate_(normalized.issueDate),
      customerName: normalized.customerName || '',
      customerAddress: normalized.customerAddress || '',
      rateItems: padItems_(normalized.rateItems),
      invoiceItems: padItems_(normalized.invoiceItems),
      displayStatus: String(source.displayStatus || '').trim(),
      systemStatus: String(source.systemStatus || '').trim(),
      cancelReason: String(source.cancelReason || '').trim(),
      cancelledAt: String(source.cancelledAt || '').trim(),
      createdAt: String(source.createdAt || '').trim(),
      updatedAt: String(source.updatedAt || '').trim(),
      links: Object.assign({}, buildAppLinks_(), source.links || {})
    },
    extra || {}
  );
}

function buildRecordDetailResponse_(row) {
  const values = getSheet_(APP_CONFIG.sheets.ledger).getRange(row, 1, 1, LEDGER_COLUMNS.lastColumn).getValues()[0];
  const issueDate = toDateOrNull_(values[LEDGER_COLUMNS.invoiceDate - 1]) || new Date();
  const paidDate = toDateOrNull_(values[LEDGER_COLUMNS.paidDate - 1]);
  const createdAt = toDateOrNull_(values[LEDGER_COLUMNS.createdAt - 1]);
  const updatedAt = toDateOrNull_(values[LEDGER_COLUMNS.updatedAt - 1]);
  const cancelledAt = toDateOrNull_(values[LEDGER_COLUMNS.cancelledAt - 1]);
  const systemStatus = String(values[LEDGER_COLUMNS.systemStatus - 1] || '').trim() || APP_CONFIG.ledger.systemStatus.active;
  const catalogMap = getCatalogMap_();
  const draft = buildDraftFromLedgerRow_(row);
  const amount = toNumber_(values[LEDGER_COLUMNS.amount - 1]);
  const recordId = String(values[LEDGER_COLUMNS.recordId - 1] || '').trim();

  return {
    row,
    recordId,
    revisionToken: buildRevisionToken_(updatedAt || createdAt),
    invoiceNumber: String(values[LEDGER_COLUMNS.rateNumber - 1] || '').trim(),
    issueDate: formatDateForUi_(issueDate),
    issueDateThai: formatThaiDate_(issueDate),
    customerName: String(values[LEDGER_COLUMNS.customerName - 1] || '').trim(),
    customerAddress: String(values[LEDGER_COLUMNS.address - 1] || '').trim(),
    amount,
    amountText: formatNumberForSheet_(amount),
    rateTotal: toNumber_(values[LEDGER_COLUMNS.rateTotal - 1]),
    invoiceTotal: toNumber_(values[LEDGER_COLUMNS.invoiceTotal - 1]),
    outstandingAmount: toNumber_(values[LEDGER_COLUMNS.outstandingAmount - 1]),
    paidDate: paidDate ? formatDateForUi_(paidDate) : '',
    displayStatus: computeDisplayStatus_(systemStatus, paidDate),
    systemStatus,
    cancelReason: String(values[LEDGER_COLUMNS.cancelReason - 1] || '').trim(),
    ratePdfUrl: String(values[LEDGER_COLUMNS.ratePdfUrl - 1] || '').trim(),
    invoicePdfUrl: String(values[LEDGER_COLUMNS.invoicePdfUrl - 1] || '').trim(),
    createdAt: createdAt ? formatDateTimeForUi_(createdAt) : '',
    updatedAt: updatedAt ? formatDateTimeForUi_(updatedAt) : '',
    cancelledAt: cancelledAt ? formatDateTimeForUi_(cancelledAt) : '',
    rateItems: enrichItemsForDisplay_(draft.rateItems, catalogMap),
    invoiceItems: enrichItemsForDisplay_(draft.invoiceItems, catalogMap),
    links: Object.assign({}, buildAppLinks_(), {
      detailUrl: buildRecordDetailUrl_(recordId),
      editUrl: buildWebAppUrl_('form', { recordId })
    })
  };
}

function enrichItemsForDisplay_(items, catalogMap) {
  return trimEmptyItems_(items).map((item) => {
    const catalogItem = catalogMap[item.code] || null;
    const unitPrice = catalogItem ? toNumber_(catalogItem.price) : 0;
    const quantity = item.quantity === '' ? 0 : toNumber_(item.quantity);
    return {
      line: item.line,
      code: item.code,
      label: catalogItem ? catalogItem.label : item.code,
      detail: item.detail,
      quantity,
      unitPrice,
      totalPrice: unitPrice * quantity
    };
  });
}

function normalizeDraftInput_(payload) {
  const source = payload || {};
  return {
    recordId: String(source.recordId || '').trim(),
    row: Number(source.row || 0) || '',
    revisionToken: String(source.revisionToken || '').trim(),
    invoiceNumber: String(source.invoiceNumber || '').trim(),
    issueDate: normalizeDateInput_(source.issueDate),
    customerName: String(source.customerName || '').trim(),
    customerAddress: String(source.customerAddress || '').trim(),
    rateItems: normalizeItems_(source.rateItems),
    invoiceItems: normalizeItems_(source.invoiceItems)
  };
}

function normalizeItems_(items) {
  const rows = Array.isArray(items) ? items : [];
  return padItems_(
    rows.map((item, index) => ({
      line: index + 1,
      code: String((item && item.code) || '').trim(),
      detail: String((item && item.detail) || '').trim(),
      quantity:
        item && item.quantity !== '' && item.quantity !== null && typeof item.quantity !== 'undefined'
          ? toNumber_(item.quantity)
          : ''
    }))
  );
}

function createEmptyItems_() {
  return padItems_([]);
}

function padItems_(items) {
  const normalized = Array.isArray(items) ? items.slice(0, APP_CONFIG.form.maxRows) : [];
  while (normalized.length < APP_CONFIG.form.maxRows) {
    normalized.push({
      line: normalized.length + 1,
      code: '',
      detail: '',
      quantity: ''
    });
  }
  return normalized.map((item, index) => ({
    line: index + 1,
    code: String(item.code || '').trim(),
    detail: String(item.detail || '').trim(),
    quantity: item.quantity === '' ? '' : toNumber_(item.quantity)
  }));
}

function trimEmptyItems_(items) {
  return (Array.isArray(items) ? items : [])
    .map((item, index) => ({
      line: index + 1,
      code: String(item.code || '').trim(),
      detail: String(item.detail || '').trim(),
      quantity: item.quantity === '' ? '' : toNumber_(item.quantity)
    }))
    .filter((item) => item.code || item.detail || item.quantity !== '');
}

function validateFormData_(draft) {
  if (!draft.customerName) {
    throw new Error('กรุณากรอกชื่อผู้ใช้บริการ');
  }
  if (!draft.customerAddress) {
    throw new Error('กรุณากรอกที่อยู่');
  }
  if (!draft.invoiceNumber) {
    throw new Error('ไม่พบเลขที่ใบแจ้งอัตราค่าบริการ');
  }

  const rateItems = trimEmptyItems_(draft.rateItems);
  const invoiceItems = trimEmptyItems_(draft.invoiceItems);
  if (!rateItems.length && !invoiceItems.length) {
    throw new Error('กรุณากรอกรายการอย่างน้อย 1 รายการ');
  }

  [['ใบแจ้งอัตราค่าบริการ', rateItems], ['ใบแจ้งค่าบริการ', invoiceItems]].forEach(([label, items]) => {
    items.forEach((item) => {
      if (!item.code) {
        throw new Error(`${label} แถวที่ ${item.line}: กรุณาเลือก Test No.`);
      }
      if (!(item.quantity > 0)) {
        throw new Error(`${label} แถวที่ ${item.line}: จำนวนต้องมากกว่า 0`);
      }
    });
  });
}

function readDraftFromBuffer_() {
  ensureSystemReady_({ refreshViews: false });
  const sheet = getSheet_(APP_CONFIG.sheets.form);
  const header = {
    recordId: sheet.getRange(APP_CONFIG.form.recordIdCell).getDisplayValue(),
    row: Number(sheet.getRange(APP_CONFIG.form.ledgerRowCell).getValue()) || '',
    invoiceNumber: sheet.getRange(APP_CONFIG.form.invoiceNumberCell).getDisplayValue(),
    issueDate: sheet.getRange(APP_CONFIG.form.rawIssueDateCell).getValue(),
    customerName: sheet.getRange(APP_CONFIG.form.customerNameCell).getDisplayValue(),
    customerAddress: sheet.getRange(APP_CONFIG.form.customerAddressCell).getDisplayValue()
  };
  return normalizeDraftInput_(
    Object.assign({}, header, {
      rateItems: extractItemsFromRange_(sheet.getRange(APP_CONFIG.form.rateRange).getValues()),
      invoiceItems: extractItemsFromRange_(sheet.getRange(APP_CONFIG.form.invoiceRange).getValues())
    })
  );
}

function writeDraftToBuffer_(payload) {
  const draft = normalizeDraftInput_(payload);
  const sheet = getSheet_(APP_CONFIG.sheets.form);
  clearBufferSheet_(sheet);

  sheet.getRange(APP_CONFIG.form.invoiceNumberCell).setNumberFormat('@').setValue(draft.invoiceNumber || '');
  sheet.getRange(APP_CONFIG.form.issueDateThaiCell).setNumberFormat('@').setValue(formatThaiDate_(draft.issueDate));
  sheet.getRange(APP_CONFIG.form.issueDateEnglishCell).setNumberFormat('@').setValue(formatEnglishDate_(draft.issueDate));
  sheet.getRange(APP_CONFIG.form.customerNameCell).setNumberFormat('@').setValue(draft.customerName || '');
  sheet.getRange(APP_CONFIG.form.customerAddressCell).setNumberFormat('@').setValue(draft.customerAddress || '');
  sheet.getRange(APP_CONFIG.form.recordIdCell).setValue(draft.recordId || '');
  sheet.getRange(APP_CONFIG.form.ledgerRowCell).setValue(draft.row || '');
  sheet.getRange(APP_CONFIG.form.rawIssueDateCell).setValue(draft.issueDate || '');

  writeItemsToRange_(sheet.getRange(APP_CONFIG.form.rateRange), draft.rateItems);
  writeItemsToRange_(sheet.getRange(APP_CONFIG.form.invoiceRange), draft.invoiceItems);
}

function clearBufferSheet_(sheet) {
  sheet.getRange(APP_CONFIG.form.rateRange).clearContent();
  sheet.getRange(APP_CONFIG.form.invoiceRange).clearContent();
  sheet.getRange(APP_CONFIG.form.invoiceNumberCell).clearContent();
  sheet.getRange(APP_CONFIG.form.issueDateThaiCell).clearContent();
  sheet.getRange(APP_CONFIG.form.issueDateEnglishCell).clearContent();
  sheet.getRange(APP_CONFIG.form.customerNameCell).clearContent();
  sheet.getRange(APP_CONFIG.form.customerAddressCell).clearContent();
  sheet.getRange(APP_CONFIG.form.recordIdCell).clearContent();
  sheet.getRange(APP_CONFIG.form.ledgerRowCell).clearContent();
  sheet.getRange(APP_CONFIG.form.rawIssueDateCell).clearContent();
}

function extractItemsFromRange_(rows) {
  return rows.map((row, index) => ({
    line: index + 1,
    code: String(row[0] || '').trim(),
    detail: String(row[1] || '').trim(),
    quantity: row[2] === '' || row[2] === null ? '' : toNumber_(row[2])
  }));
}

function writeItemsToRange_(range, items) {
  const values = padItems_(items).map((item) => [item.code || '', item.detail || '', item.quantity === '' ? '' : item.quantity]);
  range.setValues(values);
}

function updateSingleLedgerRowStatus_(sheet, row) {
  const systemStatus = String(sheet.getRange(row, LEDGER_COLUMNS.systemStatus).getValue()).trim();
  const paidDate = toDateOrNull_(sheet.getRange(row, LEDGER_COLUMNS.paidDate).getValue());
  const status = computeDisplayStatus_(systemStatus, paidDate);
  const amount = toNumber_(sheet.getRange(row, LEDGER_COLUMNS.amount).getValue());
  const outstandingAmount = status === APP_CONFIG.ledger.displayStatus.pending ? amount : 0;
  sheet.getRange(row, LEDGER_COLUMNS.receiptStatus).setValue(status);
  sheet.getRange(row, LEDGER_COLUMNS.outstandingAmount).setValue(outstandingAmount);
  formatLedgerBodyRow_(sheet, row, status);
}

function computeDisplayStatus_(systemStatus, paidDate) {
  if (String(systemStatus).trim() === APP_CONFIG.ledger.systemStatus.cancelled) {
    return APP_CONFIG.ledger.displayStatus.cancelled;
  }
  if (paidDate) {
    return APP_CONFIG.ledger.displayStatus.paid;
  }
  return APP_CONFIG.ledger.displayStatus.pending;
}

function formatLedgerBodyRow_(sheet, row, status) {
  let background = '#ffffff';
  let fontColor = '#111827';

  if (status === APP_CONFIG.ledger.displayStatus.pending) {
    background = APP_CONFIG.colors.pending;
    fontColor = APP_CONFIG.colors.pendingAccent;
  } else if (status === APP_CONFIG.ledger.displayStatus.paid) {
    background = APP_CONFIG.colors.paid;
    fontColor = APP_CONFIG.colors.paidAccent;
  } else if (status === APP_CONFIG.ledger.displayStatus.cancelled) {
    background = APP_CONFIG.colors.cancelled;
    fontColor = APP_CONFIG.colors.cancelledAccent;
  }

  sheet.getRange(row, 1, 1, 8).setBackground(background);
  sheet.getRange(row, LEDGER_COLUMNS.receiptStatus).setFontWeight('bold').setFontColor(fontColor).setHorizontalAlignment('center');
}

function findLedgerRowByRecordId_(recordId) {
  if (!recordId) {
    return 0;
  }
  const sheet = getSheet_(APP_CONFIG.sheets.ledger);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return 0;
  }
  const values = sheet.getRange(2, LEDGER_COLUMNS.recordId, lastRow - 1, 1).getDisplayValues();
  const index = values.findIndex((row) => String(row[0]).trim() === String(recordId).trim());
  return index >= 0 ? index + 2 : 0;
}

function findLedgerRowByInvoiceNumber_(invoiceNumber) {
  if (!invoiceNumber) {
    return 0;
  }
  const sheet = getSheet_(APP_CONFIG.sheets.ledger);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return 0;
  }
  const values = sheet.getRange(2, LEDGER_COLUMNS.rateNumber, lastRow - 1, 1).getDisplayValues();
  const index = values.findIndex((row) => String(row[0]).trim() === String(invoiceNumber).trim());
  return index >= 0 ? index + 2 : 0;
}

function getRecordIdFromSelection_() {
  const activeSheet = getSpreadsheet_().getActiveSheet();
  const activeRange = activeSheet.getActiveRange();
  if (!activeRange || activeRange.getRow() <= 1) {
    throw new Error('กรุณาเลือกแถวข้อมูลก่อน');
  }

  if (activeSheet.getName() === APP_CONFIG.sheets.ledger) {
    return String(activeSheet.getRange(activeRange.getRow(), LEDGER_COLUMNS.recordId).getDisplayValue()).trim();
  }

  throw new Error('กรุณาเลือกแถวจากชีต "ทะเบียนลูกหนี้"');
}

function generateNextInvoiceNumber_(issueDate) {
  const safeDate = normalizeDateInput_(issueDate);
  const buddhistYear = safeDate.getFullYear() + 543;
  const yearSuffix = String(buddhistYear).slice(-2);
  const sheet = getSheet_(APP_CONFIG.sheets.ledger);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return `001/${yearSuffix}`;
  }

  const values = sheet.getRange(2, LEDGER_COLUMNS.rateNumber, lastRow - 1, 1).getDisplayValues().flat();
  const maxNumber = values.reduce((currentMax, value) => {
    const match = String(value).trim().match(/^(\d+)\/(\d{2})$/);
    if (!match || match[2] !== yearSuffix) {
      return currentMax;
    }
    return Math.max(currentMax, Number(match[1]));
  }, 0);

  return `${String(maxNumber + 1).padStart(3, '0')}/${yearSuffix}`;
}

function buildRevisionToken_(dateValue) {
  const date = toDateOrNull_(dateValue);
  return date ? String(date.getTime()) : '';
}

function exportDocumentForLedgerRow_(context, type) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const draft = buildDraftFromLedgerRow_(context.row);
    writeDraftToBuffer_(draft);
    SpreadsheetApp.flush();

    const sheetName = type === 'rateStatement' ? APP_CONFIG.sheets.rateStatement : APP_CONFIG.sheets.serviceInvoice;
    const title = type === 'rateStatement' ? 'ใบแจ้งอัตราค่าบริการ' : 'ใบแจ้งค่าบริการ';
    const existingUrl = type === 'rateStatement' ? context.ratePdfUrl : context.invoicePdfUrl;
    const fileName = [
      sanitizeFileName_(context.invoiceNumber),
      title,
      sanitizeFileName_(context.customerName || 'ไม่ระบุชื่อ')
    ].join('_') + '.pdf';

    return exportSheetToPdf_(sheetName, fileName, existingUrl);
  } finally {
    lock.releaseLock();
  }
}

function updateLedgerDocumentLinks_(row, links) {
  const ledgerSheet = getSheet_(APP_CONFIG.sheets.ledger);
  const updates = [];
  if (Object.prototype.hasOwnProperty.call(links, 'ratePdfUrl')) {
    updates.push({ column: LEDGER_COLUMNS.ratePdfUrl, value: links.ratePdfUrl || '' });
  }
  if (Object.prototype.hasOwnProperty.call(links, 'invoicePdfUrl')) {
    updates.push({ column: LEDGER_COLUMNS.invoicePdfUrl, value: links.invoicePdfUrl || '' });
  }
  updates.push({ column: LEDGER_COLUMNS.updatedAt, value: new Date() });
  updates.forEach((item) => ledgerSheet.getRange(row, item.column).setValue(item.value));
}

function exportSheetToPdf_(sheetName, fileName, existingFileUrl) {
  const spreadsheet = getSpreadsheet_();
  const sheet = getSheet_(sheetName);
  SpreadsheetApp.flush();
  Utilities.sleep(350);

  const url = buildSheetPdfExportUrl_(spreadsheet.getId(), sheet.getSheetId());
  const blob = fetchSheetPdfBlobWithRetry_(url, sheetName).setName(fileName);
  return upsertPdfFile_(blob, fileName, existingFileUrl);
}

function buildSheetPdfExportUrl_(spreadsheetId, sheetId) {
  const params = Object.keys(APP_CONFIG.pdfExport)
    .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(APP_CONFIG.pdfExport[key])}`)
    .join('&');
  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${params}&gid=${sheetId}`;
}

function fetchSheetPdfBlobWithRetry_(url, sheetName) {
  const maxAttempts = Number(APP_CONFIG.pdfRetry.maxAttempts || 4);
  const initialDelayMs = Number(APP_CONFIG.pdfRetry.initialDelayMs || 900);
  let lastResponse = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      return response.getBlob();
    }

    lastResponse = response;
    const shouldRetry = responseCode === 429 || responseCode >= 500;
    if (!shouldRetry || attempt === maxAttempts) {
      break;
    }

    SpreadsheetApp.flush();
    Utilities.sleep(initialDelayMs * attempt);
  }

  const responseCode = lastResponse ? lastResponse.getResponseCode() : 'UNKNOWN';
  const responseText = lastResponse ? summarizeFetchError_(lastResponse.getContentText()) : '';
  const detailText = responseText ? `: ${responseText}` : '';
  throw new Error(`สร้าง PDF ชีต "${sheetName}" ไม่สำเร็จ (${responseCode})${detailText}`);
}

function summarizeFetchError_(text) {
  return String(text || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 180);
}

function upsertPdfFile_(blob, fileName, existingFileUrl) {
  const existingFileId = extractDriveFileId_(existingFileUrl);
  if (existingFileId) {
    return replacePdfFileContent_(existingFileId, blob, fileName);
  }

  const folder = getOrCreateExportFolder_();
  const file = folder.createFile(blob.setName(fileName));
  return {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName()
  };
}

function replacePdfFileContent_(fileId, blob, fileName) {
  const url = `https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(fileId)}?uploadType=media&supportsAllDrives=true`;
  const response = UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/pdf',
    payload: blob.getBytes(),
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`อัปเดตไฟล์ PDF เดิมไม่สำเร็จ (${response.getResponseCode()})`);
  }

  const file = DriveApp.getFileById(fileId);
  if (fileName && typeof file.setName === 'function') {
    file.setName(fileName);
  }
  return {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName()
  };
}

function extractDriveFileId_(urlOrId) {
  const input = String(urlOrId || '').trim();
  if (!input) {
    return '';
  }

  if (/^[a-zA-Z0-9_-]{20,}$/.test(input)) {
    return input;
  }

  const pathMatch = input.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (pathMatch && pathMatch[1]) {
    return pathMatch[1];
  }

  const queryMatch = input.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return queryMatch && queryMatch[1] ? queryMatch[1] : '';
}

function getOrCreateExportFolder_() {
  const spreadsheet = getSpreadsheet_();
  const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  const folderName = `${spreadsheet.getName()}_ไฟล์ใบแจ้งหนี้`;
  const parents = spreadsheetFile.getParents();

  if (parents.hasNext()) {
    const parent = parents.next();
    const existing = parent.getFoldersByName(folderName);
    return existing.hasNext() ? existing.next() : parent.createFolder(folderName);
  }

  const existingRootFolders = DriveApp.getFoldersByName(folderName);
  return existingRootFolders.hasNext() ? existingRootFolders.next() : DriveApp.createFolder(folderName);
}

function applyInvoiceCancellation_(row, reason) {
  const ledgerSheet = getSheet_(APP_CONFIG.sheets.ledger);
  ledgerSheet.getRange(row, LEDGER_COLUMNS.systemStatus).setValue(APP_CONFIG.ledger.systemStatus.cancelled);
  ledgerSheet.getRange(row, LEDGER_COLUMNS.cancelledAt).setValue(new Date());
  ledgerSheet.getRange(row, LEDGER_COLUMNS.cancelReason).setValue(reason || '');
  ledgerSheet.getRange(row, LEDGER_COLUMNS.updatedAt).setValue(new Date());
  updateSingleLedgerRowStatus_(ledgerSheet, row);
  refreshAllViews_();
}

function getCatalogOptions_() {
  const sheet = getSheet_(APP_CONFIG.sheets.catalog);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, 6)
    .getValues()
    .map((row) => ({
      code: String(row[0] || '').trim(),
      name: String(row[1] || '').trim(),
      collectPrice: toNumber_(row[2]),
      analysisPrice: toNumber_(row[3]),
      totalPrice: toNumber_(row[4]),
      note: String(row[5] || '').trim()
    }))
    .filter((item) => item.code)
    .map((item) =>
      Object.assign({}, item, {
        label: `${item.code} - ${item.name}`,
        price: item.totalPrice
      })
    );
}

function getCatalogMap_() {
  return getCatalogOptions_().reduce((result, item) => {
    result[item.code] = item;
    return result;
  }, {});
}

function buildAppLinks_() {
  return {
    ledgerUrl: buildSheetUrlByName_(APP_CONFIG.sheets.ledger),
    dashboardUrl: buildWebAppUrl_('dashboard'),
    detailUrl: buildWebAppUrl_('detail'),
    formUrl: buildWebAppUrl_('form'),
    apiUrl: buildApiUrl_(),
    dashboardApiUrl: buildApiUrl_('dashboard'),
    formApiUrl: buildApiUrl_('formBootstrap'),
    detailApiUrl: buildApiUrl_('detail'),
    saveApiUrl: buildApiUrl_('save'),
    webAppUrl: ScriptApp.getService().getUrl() || ''
  };
}

function buildRecordId_(invoiceNumber) {
  const stamp = Utilities.formatDate(new Date(), APP_CONFIG.timezone, 'yyyyMMddHHmmss');
  return `REC-${sanitizeFileName_(invoiceNumber)}-${stamp}`;
}

function mapRecordForUi_(record) {
  return {
    recordId: record.recordId,
    customerName: record.customerName,
    invoiceNumber: record.invoiceNumber,
    amount: record.amount,
    outstandingAmount: record.outstandingAmount,
    paidAmount: record.paidAmount,
    displayStatus: record.displayStatus,
    invoiceDate: formatDateForUi_(record.invoiceDate),
    paidDate: formatDateForUi_(record.paidDate),
    ratePdfUrl: record.ratePdfUrl,
    invoicePdfUrl: record.invoicePdfUrl
  };
}

function parseJsonArray_(value) {
  if (!value) {
    return [];
  }
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    console.error(error);
    return [];
  }
}

function formatThaiDate_(date) {
  const safeDate = normalizeDateInput_(date);
  const months = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  return `${safeDate.getDate()} ${months[safeDate.getMonth()]} พ.ศ. ${safeDate.getFullYear() + 543}`;
}

function formatEnglishDate_(date) {
  const safeDate = normalizeDateInput_(date);
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  return `${safeDate.getDate()} ${months[safeDate.getMonth()]} ${safeDate.getFullYear()}`;
}

function formatDateForUi_(date) {
  if (!date) {
    return '';
  }
  return Utilities.formatDate(normalizeDateInput_(date), APP_CONFIG.timezone, 'dd/MM/yyyy');
}

function formatDateTimeForUi_(date) {
  return Utilities.formatDate(normalizeDateInput_(date), APP_CONFIG.timezone, 'dd/MM/yyyy HH:mm');
}

function formatMonthLabelShort_(date) {
  const safeDate = normalizeDateInput_(date);
  const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  return `${months[safeDate.getMonth()]} ${String((safeDate.getFullYear() + 543)).slice(-2)}`;
}

function formatDateInputValue_(date) {
  return Utilities.formatDate(normalizeDateInput_(date), APP_CONFIG.timezone, 'yyyy-MM-dd');
}

function formatNumberForSheet_(value) {
  if (typeof value === 'number') {
    return value.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return String(value);
}

function styleTableHeader_(range) {
  range.setBackground(APP_CONFIG.colors.brand).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
}

function toDateOrNull_(value) {
  if (!value) {
    return null;
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return value;
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function normalizeDateInput_(value) {
  return toDateOrNull_(value) || new Date();
}

function toNumber_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }
  if (value === '' || value === null || typeof value === 'undefined') {
    return 0;
  }
  const parsed = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

function sumBy_(records, key) {
  return records.reduce((total, record) => total + toNumber_(record[key]), 0);
}

function roundToTwo_(value) {
  return Math.round(toNumber_(value) * 100) / 100;
}

function dateValue_(date) {
  return date ? normalizeDateInput_(date).getTime() : 0;
}

function sanitizeFileName_(value) {
  return String(value || '')
    .trim()
    .replace(/[\\/:*?"<>|#%&{}$!'@+=`~]/g, '-')
    .replace(/\s+/g, ' ')
    .substring(0, 80);
}

function buildSheetUrlByName_(sheetName) {
  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(sheetName);
  return sheet ? `${spreadsheet.getUrl()}#gid=${sheet.getSheetId()}` : spreadsheet.getUrl();
}

function buildRecordDetailUrl_(recordId) {
  return buildWebAppUrl_('detail', { recordId: String(recordId || '').trim() });
}

function buildWebAppUrl_(page, extraParams) {
  const baseUrl = ScriptApp.getService().getUrl() || '';
  if (!baseUrl) {
    return '';
  }

  const params = Object.assign({ page }, extraParams || {});
  const query = Object.keys(params)
    .filter((key) => params[key] !== '' && params[key] !== null && typeof params[key] !== 'undefined')
    .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(String(params[key]))}`)
    .join('&');

  return query ? `${baseUrl}?${query}` : baseUrl;
}

function buildApiUrl_(action, extraParams) {
  const baseUrl = ScriptApp.getService().getUrl() || '';
  if (!baseUrl) {
    return '';
  }

  const params = Object.assign({ mode: 'api' }, action ? { action } : {}, extraParams || {});
  const query = Object.keys(params)
    .filter((key) => params[key] !== '' && params[key] !== null && typeof params[key] !== 'undefined')
    .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(String(params[key]))}`)
    .join('&');

  return query ? `${baseUrl}?${query}` : baseUrl;
}

function getApiAction_(e) {
  return String((e && e.parameter && e.parameter.action) || 'health').trim().toLowerCase();
}

function parseApiBody_(e) {
  if (!e) {
    return {};
  }

  const raw = e.postData && typeof e.postData.contents === 'string' ? e.postData.contents.trim() : '';
  if (raw) {
    try {
      return JSON.parse(raw);
    } catch (error) {
      console.error(error);
    }
  }

  const params = Object.assign({}, (e && e.parameter) || {});
  if (params.payload) {
    try {
      return JSON.parse(String(params.payload));
    } catch (error) {
      console.error(error);
    }
  }

  return params;
}

function getApiResponseMode_(e, body) {
  return String(
    (e && e.parameter && e.parameter.responseMode) ||
      (body && body.responseMode) ||
      ''
  )
    .trim()
    .toLowerCase();
}

function createApiSuccessOutput_(data, e) {
  const payload = {
    ok: true,
    data,
    timestamp: new Date().toISOString()
  };
  return createApiOutput_(payload, e);
}

function createApiErrorOutput_(error, e) {
  const payload = {
    ok: false,
    error: {
      message: error && error.message ? error.message : String(error || 'Unknown API error')
    },
    timestamp: new Date().toISOString()
  };
  return createApiOutput_(payload, e);
}

function createApiOutput_(payload, e) {
  const callback = String((e && e.parameter && e.parameter.callback) || '').trim();
  if (callback) {
    const safeCallback = callback.replace(/[^\w.$]/g, '');
    if (!safeCallback) {
      const fallback = ContentService.createTextOutput(
        JSON.stringify({
          ok: false,
          error: {
            message: 'callback ไม่ถูกต้อง'
          },
          timestamp: new Date().toISOString()
        })
      );
      return fallback.setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(`${safeCallback}(${JSON.stringify(payload)})`).setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );
  }

  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function createApiRedirectOutput_(action, data, e, body) {
  const targetUrl = buildRedirectTargetUrl_(action, data, e, body);
  return createRedirectHtmlOutput_(targetUrl, 'กำลังเปิดหน้ารายละเอียดเอกสาร');
}

function createApiRedirectErrorOutput_(error, e, body) {
  const fallbackUrl = appendQueryParams_(
    String(
      (e && e.parameter && e.parameter.errorTo) ||
        (e && e.parameter && e.parameter.returnTo) ||
        (body && body.errorTo) ||
        (body && body.returnTo) ||
        ''
    ).trim(),
    {
      apiError: error && error.message ? error.message : String(error || 'Unknown API error'),
      apiBase: buildApiUrl_()
    }
  );

  return createRedirectHtmlOutput_(fallbackUrl, 'ไม่สามารถบันทึกข้อมูลได้');
}

function buildRedirectTargetUrl_(action, data, e, body) {
  const returnTo = String(
    (e && e.parameter && e.parameter.returnTo) ||
      (body && body.returnTo) ||
      ''
  ).trim();

  if (action === 'save' && returnTo && data && data.recordId) {
    return appendQueryParams_(returnTo, {
      recordId: data.recordId,
      apiBase: buildApiUrl_()
    });
  }

  return appendQueryParams_(returnTo, {
    apiBase: buildApiUrl_()
  });
}

function createRedirectHtmlOutput_(targetUrl, title) {
  const safeTitle = JSON.stringify(String(title || 'กำลังนำทาง'));
  const safeTarget = JSON.stringify(String(targetUrl || ''));
  const html = `<!DOCTYPE html>
<html lang="th">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${JSON.parse(safeTitle)}</title>
    <style>
      body {
        margin: 0;
        min-height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
        background: linear-gradient(180deg, #f8f5ef 0%, #eef6f5 100%);
        color: #1f2937;
        font-family: Arial, sans-serif;
      }
      .box {
        padding: 24px 28px;
        border-radius: 20px;
        background: rgba(255, 255, 255, 0.92);
        box-shadow: 0 18px 36px rgba(15, 23, 42, 0.08);
        text-align: center;
      }
    </style>
  </head>
  <body>
    <div class="box">
      <div id="message">${JSON.parse(safeTitle)}</div>
    </div>
    <script>
      const targetUrl = ${safeTarget};
      if (targetUrl) {
        window.top.location.replace(targetUrl);
      }
    </script>
  </body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle(String(title || 'กำลังนำทาง'))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function appendQueryParams_(url, params) {
  const rawUrl = String(url || '').trim();
  if (!rawUrl) {
    return '';
  }

  const hashIndex = rawUrl.indexOf('#');
  const base = hashIndex >= 0 ? rawUrl.substring(0, hashIndex) : rawUrl;
  const hash = hashIndex >= 0 ? rawUrl.substring(hashIndex) : '';
  const query = Object.keys(params || {})
    .filter((key) => params[key] !== '' && params[key] !== null && typeof params[key] !== 'undefined')
    .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(String(params[key]))}`)
    .join('&');

  if (!query) {
    return `${base}${hash}`;
  }

  return `${base}${base.includes('?') ? '&' : '?'}${query}${hash}`;
}

function safeToast_(message, title, seconds) {
  try {
    getSpreadsheet_().toast(message, title || 'ระบบลูกหนี้', seconds || 5);
  } catch (error) {
    console.error(error);
  }
}

function sheetRef_(sheetName, a1) {
  return `'${String(sheetName).replace(/'/g, "''")}'!${a1}`;
}
