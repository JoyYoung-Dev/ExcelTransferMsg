
<script setup>
import { computed, onMounted, ref, watch } from 'vue';
import * as XLSX from 'xlsx';

const emit = defineEmits(['result-generated']);

const ACCEPTED_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel'
];
const ACCEPTED_EXTENSIONS = ['.xlsx', '.xls'];
const STORAGE_KEY = 'excel-reader-cache';
const MAX_CACHE_ROWS_PER_SHEET = 500;
const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB is a safe browser limit
const MAX_H_ROWS = 100;

const dragActive = ref(false);
const dragDepth = ref(0);
const status = ref('idle');
const errorMessage = ref('');
const fileInfo = ref(null);
const sheets = ref([]);
const selectedSheet = ref('');
const searchTerm = ref('');
const lastUpdated = ref(null);
const fileInput = ref(null);

const isBusy = computed(() => status.value === 'loading');
const hasData = computed(() => sheets.value.length > 0);

const statusChip = computed(() => {
  switch (status.value) {
    case 'loading':
      return { label: '解析中', tone: 'warning' };
    case 'ready':
      return { label: '已就绪', tone: 'success' };
    case 'error':
      return { label: '发生错误', tone: 'danger' };
    default:
      return { label: '等待上传', tone: 'info' };
  }
});

const sheetOptions = computed(() =>
  sheets.value.map((sheet) => ({
    label: `${sheet.name}（${Math.max(sheet.rows.length - 1, 0)} 行）`,
    value: sheet.name
  }))
);

const currentSheet = computed(() =>
  sheets.value.find((sheet) => sheet.name === selectedSheet.value)
);

const currentRows = computed(() => currentSheet.value?.rows ?? []);
const headers = computed(() => currentRows.value[0] ?? []);
const normalizedHeaders = computed(() =>
  headers.value.length
    ? headers.value.map((header, index) =>
        header ? String(header) : `列 ${index + 1}`
      )
    : []
);
const bodyRows = computed(() =>
  currentRows.value.length > 1 ? currentRows.value.slice(1) : []
);
const filteredRows = computed(() => {
  if (!searchTerm.value) {
    return bodyRows.value;
  }
  const keyword = searchTerm.value.trim().toLowerCase();
  if (!keyword) {
    return bodyRows.value;
  }
  return bodyRows.value.filter((row) =>
    row.some((cell) => String(cell ?? '').toLowerCase().includes(keyword))
  );
});

const summary = computed(() => ({
  totalRows: bodyRows.value.length,
  filteredRows: filteredRows.value.length,
  columns: normalizedHeaders.value.length
}));

const dragHint = computed(() =>
  isBusy.value ? '正在解析文件，请稍候…' : '拖拽文件到此区域，或点击按钮从设备选择文件。'
);

const lastUpdatedText = computed(() =>
  lastUpdated.value ? lastUpdated.value.toLocaleString() : ''
);

function triggerFileDialog() {
  fileInput.value?.click();
}

function onFileChange(event) {
  const [file] = event.target.files ?? [];
  if (file) {
    handleIncomingFile(file);
  }
  event.target.value = '';
}

function onDragEnter(event) {
  event.preventDefault();
  if (isBusy.value) return;
  dragDepth.value += 1;
  dragActive.value = true;
}

function onDragOver(event) {
  event.preventDefault();
  event.dataTransfer.dropEffect = isBusy.value ? 'none' : 'copy';
}

function onDragLeave(event) {
  event.preventDefault();
  dragDepth.value = Math.max(dragDepth.value - 1, 0);
  if (!dragDepth.value) {
    dragActive.value = false;
  }
}

function onDrop(event) {
  event.preventDefault();
  dragDepth.value = 0;
  dragActive.value = false;
  const files = event.dataTransfer?.files;
  if (files?.length) {
    handleIncomingFile(files[0]);
  }
}

function handleIncomingFile(file) {
  if (isBusy.value) return;
  if (!isExcelFile(file)) {
    return setError('仅支持 .xlsx 或 .xls 文件。');
  }
  if (file.size > MAX_FILE_SIZE) {
    return setError('文件体积超过 10 MB，请先压缩或拆分后再试。');
  }
  errorMessage.value = '';
  searchTerm.value = '';
  processFile(file);
}

function isExcelFile(file) {
  const extension = (file.name?.toLowerCase().match(/(\.[^.]+)$/)?.[1]) ?? '';
  return ACCEPTED_EXTENSIONS.includes(extension) || ACCEPTED_TYPES.includes(file.type);
}

function processFile(file) {
  status.value = 'loading';
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const buffer = event.target?.result;

      const workbook = XLSX.read(buffer, { type: 'array' });
      const workbookSheets = workbook.SheetNames.map((name) => ({
        name,
        rows: extractSheetRows(workbook.Sheets[name])
      }));

      if (!workbookSheets.some((sheet) => sheet.rows.length)) {
        throw new Error('EMPTY_WORKBOOK');
      }

      sheets.value = workbookSheets;
      const initialSheet =
        workbookSheets.find((sheet) => sheet.rows.length) ?? workbookSheets[0];
      selectedSheet.value = initialSheet?.name ?? '';
      fileInfo.value = {
        name: file.name,
        size: file.size,
        sizeLabel: formatBytes(file.size),
        sheetCount: workbookSheets.length,
        lastModifiedLabel: new Date(file.lastModified).toLocaleString()
      };
      lastUpdated.value = new Date();
      status.value = 'ready';
      persistState();
      emit('result-generated', buildResultText(fileInfo.value, workbookSheets));
    } catch (error) {
      console.error(error);
      setError('解析 Excel 文件时出现问题，请检查文件后重试。');
    }
  };
  reader.onerror = () => {
    setError('读取文件失败，请重新上传。');
  };
  reader.readAsArrayBuffer(file);
}


// Limit JSON conversion to the actual used cells to avoid iterating millions of blanks.
function extractSheetRows(sheet) {
  if (!sheet) {
    return [];
  }
  const bounds = getSheetValueBounds(sheet);
  if (!bounds) {
    return [];
  }

  const jsonOptions = {
    header: 1,
    blankrows: true,
    defval: '',
    range: XLSX.utils.encode_range(bounds)
  };
  const rows = XLSX.utils.sheet_to_json(sheet, jsonOptions);
  return trimTrailingEmptyRows(rows);
}

function getSheetValueBounds(sheet) {
  if (!sheet) {
    return null;
  }
  const cellAddresses = Object.keys(sheet).filter((address) => !address.startsWith('!'));
  if (!cellAddresses.length) {
    return null;
  }

  let minRow = Infinity;
  let maxRow = -1;
  let minCol = Infinity;
  let maxCol = -1;
  let hasContent = false;

  cellAddresses.forEach((address) => {
    const cell = sheet[address];
    const value = cell?.v ?? cell?.w;
    if (!hasMeaningfulCellValue(value)) {
      return;
    }
    hasContent = true;
    const { r, c } = XLSX.utils.decode_cell(address);
    if (r < minRow) minRow = r;
    if (r > maxRow) maxRow = r;
    if (c < minCol) minCol = c;
    if (c > maxCol) maxCol = c;
  });

  if (!hasContent || !Number.isFinite(minRow) || !Number.isFinite(minCol)) {
    return null;
  }

  return {
    s: { r: minRow, c: minCol },
    e: { r: Math.max(maxRow, minRow), c: Math.max(maxCol, minCol) }
  };
}

function hasMeaningfulCellValue(value) {
  if (value === null || value === undefined) {
    return false;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value);
  }
  if (typeof value === 'boolean') {
    return true;
  }
  return String(value).trim() !== '';
}

function trimTrailingEmptyRows(rows) {
  let endIndex = rows.length - 1;
  while (endIndex >= 0 && !rowHasMeaningfulData(rows[endIndex])) {
    endIndex -= 1;
  }
  return endIndex >= 0 ? rows.slice(0, endIndex + 1) : [];
}

function rowHasMeaningfulData(row) {
  if (!Array.isArray(row)) {
    return false;
  }
  return row.some((cell) => hasMeaningfulCellValue(cell));
}

function setError(message) {
  errorMessage.value = message;
  status.value = 'error';
}

function resetAll() {
  fileInfo.value = null;
  sheets.value = [];
  selectedSheet.value = '';
  searchTerm.value = '';
  status.value = 'idle';
  errorMessage.value = '';
  lastUpdated.value = null;
  clearCache();
  emit('result-generated', '');
}

function formatBytes(bytes) {
  if (!bytes && bytes !== 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let size = bytes;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }
  const value = size >= 10 || unitIndex === 0 ? Math.round(size) : size.toFixed(1);
  return `${value} ${units[unitIndex]}`;
}

function buildResultText(info, workbookSheets) {
  if (!workbookSheets?.length) {
    return info ? `未能从 ${info.name} 提取到数据。` : '';
  }

  const sections = workbookSheets
    .map((sheet) => ({
      name: sheet.name,
      content: buildSectionFromSheet(sheet)
    }))
    .filter((entry) => entry.content);

  if (sections.length) {
    const shouldShowSheetLabel = sections.length > 1 || workbookSheets.length > 1;
    return sections
      .map((entry) => (shouldShowSheetLabel ? `【${entry.name}】
${entry.content}` : entry.content))
      .join('\n');
  }

  return info ? `未能从 ${info.name} 解析出符合规则的内容。` : '';
}



function buildSectionFromSheet(sheet) {
  const rows = (sheet.rows ?? []).slice(0, MAX_H_ROWS);
  const H_INDEX = 7;
  const I_INDEX = 8;

  if (!rows.length) {
    return '';
  }

  const hValues = rows.map((row) => sanitizeText(row?.[H_INDEX]));
  const firstRowIndex = hValues.findIndex(Boolean);
  const lastHValueRow = findLastIndex(hValues, Boolean);
  const maxRangeRow = Math.min(rows.length - 1, MAX_H_ROWS - 1);
  const lastRowIndex = Math.max(lastHValueRow, maxRangeRow);

  if (firstRowIndex === -1 || lastRowIndex === -1) {
    return '';
  }

  const sections = [];
  let cursor = firstRowIndex;

  while (cursor <= lastRowIndex) {
    const dateRowIndex = findNextDateRow(rows, cursor, lastRowIndex, H_INDEX);
    if (dateRowIndex === -1) {
      break;
    }

    const block = buildBlockFromDate(rows, dateRowIndex, lastRowIndex, H_INDEX, I_INDEX);
    if (block.content) {
      sections.push(block.content);
    }
    cursor = Math.max(block.nextRow ?? dateRowIndex + 1, dateRowIndex + 1);
  }

  if (!sections.length) {
    return '';
  }

  const rangeLabel = `H列范围：H${firstRowIndex + 1} - H${lastRowIndex + 1}`;
  return [rangeLabel, ...sections].join('\n').trim();
}

function findNextDateRow(rows, start, end, hIndex) {
  for (let rowIndex = start; rowIndex <= end; rowIndex += 1) {
    if (looksLikeDateCell(rows[rowIndex]?.[hIndex])) {
      return rowIndex;
    }
  }
  return -1;
}



function buildBlockFromDate(rows, dateRowIndex, lastRowIndex, H_INDEX, I_INDEX) {
  const dateLabel = formatDateLabel(rows[dateRowIndex]?.[H_INDEX]);
  if (!dateLabel) {
    return { content: '', nextRow: dateRowIndex + 1 };
  }

  const headerRowIndex = findHeaderRowIndex(rows, dateRowIndex, I_INDEX);
  const stores = extractStores(rows, headerRowIndex, I_INDEX);
  if (!stores.length) {
    return { content: dateLabel, nextRow: dateRowIndex + 1 };
  }

  let hasStartedProducts = false;
  let nextRowPointer = dateRowIndex + 1;

  for (let rowIndex = dateRowIndex; rowIndex <= lastRowIndex; rowIndex += 1) {
    const cellValue = rows[rowIndex]?.[H_INDEX];
    const isDateCell = looksLikeDateCell(cellValue);

    if (isDateCell && rowIndex !== dateRowIndex) {
      nextRowPointer = rowIndex;
      break;
    }

    const productName = resolveProductName(rows[rowIndex], H_INDEX);
    if (!productName) {
      nextRowPointer = rowIndex + 1;
      continue;
    }

    hasStartedProducts = true;
    nextRowPointer = rowIndex + 1;

    stores.forEach((store) => {
      const quantity = parseQuantityValue(rows[rowIndex]?.[store.columnIndex]);
      if (quantity !== null) {
        store.items.push({ name: productName, quantity });
      }
    });
  }

  if (nextRowPointer <= dateRowIndex + 1) {
    nextRowPointer = lastRowIndex + 1;
  }

  const storeSections = stores
    .filter((store) => store.items.length)
    .map((store) => {
      const entries = store.items
        .map((item) => `- ${item.name}：${formatQuantity(item.quantity)} pcs`)
        .join('\n');
      return `${store.name}\n${entries}`;
    });

  const content = [dateLabel, ...storeSections].join('\n').trim();

  return { content, nextRow: nextRowPointer };
}

function resolveProductName(row, H_INDEX) {
  if (!row) {
    return '';
  }
  const hValue = row[H_INDEX];
  if (hValue && !looksLikeDateCell(hValue)) {
    const normalized = sanitizeProductName(hValue);
    if (normalized) {
      return normalized;
    }
  }

  const fallbackColumns = [H_INDEX - 1, 1];
  for (const columnIndex of fallbackColumns) {
    if (columnIndex < 0) {
      continue;
    }
    const fallbackValue = sanitizeProductName(row[columnIndex]);
    if (fallbackValue) {
      return fallbackValue;
    }
  }

  return '';
}

function findHeaderRowIndex(rows, dateRowIndex, I_INDEX) {
  if (dateRowIndex <= 0) {
    return -1;
  }
  for (let rowIndex = dateRowIndex - 1; rowIndex >= 0; rowIndex -= 1) {
    const candidateRow = rows[rowIndex] ?? [];
    const hasStoreLabel = candidateRow.slice(I_INDEX).some((cell) => sanitizeHeaderCell(cell));
    if (hasStoreLabel) {
      return rowIndex;
    }
  }
  return dateRowIndex - 1;
}

function extractStores(rows, headerRowIndex, I_INDEX) {
  if (headerRowIndex < 0) {
    return [];
  }
  const headerRow = rows[headerRowIndex] ?? [];
  const stores = [];

  for (let colIndex = I_INDEX; colIndex < headerRow.length; colIndex += 1) {
    const label = sanitizeHeaderCell(headerRow[colIndex]);
    if (!label) {
      break;
    }
    stores.push({ name: label, columnIndex: colIndex, items: [] });
  }

  return stores;
}

function findLastIndex(array, predicate) {
  for (let index = array.length - 1; index >= 0; index -= 1) {
    if (predicate(array[index])) {
      return index;
    }
  }
  return -1;
}

function looksLikeDateCell(value) {
  if (value === null || value === undefined) {
    return false;
  }
  if (value instanceof Date) {
    return true;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value > 10000;
  }
  const text = sanitizeText(value);
  if (!text) {
    return false;
  }
  const datePattern = /(\d{1,4}[\/-]\d{1,2}[\/-]\d{1,4})/;
  const weekdayPattern = /星期[一二三四五六日天]/;
  return datePattern.test(text) || weekdayPattern.test(text);
}

function formatDateLabel(value) {
  if (value instanceof Date) {
    return formatDateForDisplay(value);
  }
  if (typeof value === 'number' && Number.isFinite(value) && XLSX.SSF?.parse_date_code) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const date = new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S);
      return formatDateForDisplay(date);
    }
  }
  const text = sanitizeText(value);
  if (!text) {
    return '';
  }
  return text.replace(/[\s　]+/g, ' ').trim();
}

function formatDateForDisplay(date) {
  return date.toLocaleDateString('zh-CN', { year: 'numeric', month: '2-digit', day: '2-digit', weekday: 'long' });
}

function sanitizeHeaderCell(value) {
  return sanitizeText(value);
}

function sanitizeProductName(value) {
  return sanitizeText(value);
}

function sanitizeText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).replace(/[\r\n\t]+/g, '').trim();
}

function parseQuantityValue(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value > 0 ? value : null;
  }
  const text = String(value).replace(/[\s,]/g, '');
  if (!text) {
    return null;
  }
  const parsed = Number(text);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return null;
  }
  return parsed;
}

function formatQuantity(value) {
  if (Number.isInteger(value)) {
    return `${value}`;
  }
  return `${Number(value.toFixed(2))}`;
}

function formatCell(value) {
  if (value === null || value === undefined || value === '') {
    return '—';
  }
  return String(value);
}

function persistState() {
  if (!hasData.value || !canUseStorage()) {
    return clearCache();
  }
  try {
    const payload = {
      fileInfo: fileInfo.value,
      sheets: sheets.value.map((sheet) => ({
        name: sheet.name,
        rows: sheet.rows.slice(0, MAX_CACHE_ROWS_PER_SHEET)
      })),
      selectedSheet: selectedSheet.value,
      timestamp: lastUpdated.value?.toISOString() ?? new Date().toISOString()
    };
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  } catch (error) {
    console.warn('无法写入本地缓存', error);
  }
}

function clearCache() {
  if (!canUseStorage()) return;
  window.localStorage.removeItem(STORAGE_KEY);
}

function canUseStorage() {
  return typeof window !== 'undefined' && 'localStorage' in window;
}

function restoreFromStorage() {
  if (!canUseStorage()) return;
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const payload = JSON.parse(raw);
    if (!payload?.sheets?.length) return;
    fileInfo.value = payload.fileInfo ?? null;
    sheets.value = payload.sheets;
    selectedSheet.value = payload.selectedSheet ?? payload.sheets[0].name;
    status.value = 'ready';
    searchTerm.value = '';
    lastUpdated.value = payload.timestamp ? new Date(payload.timestamp) : null;
    emit('result-generated', buildResultText(fileInfo.value, sheets.value));
  } catch (error) {
    console.warn('读取本地缓存失败', error);
  }
}

watch(selectedSheet, () => {
  if (hasData.value) {
    persistState();
  }
});

onMounted(() => {
  restoreFromStorage();
});
</script>

<template>
  <section class="reader-card" aria-live="polite">
    <header class="reader-card__header">
      <div>
        <p class="eyebrow">数据读取</p>
        <h2>上传并解析 Excel</h2>
        <p>支持 .xlsx / .xls，整个过程完全在浏览器中完成。</p>
      </div>
      <span class="status-chip" :class="`status-chip--${statusChip.tone}`">
        {{ statusChip.label }}
      </span>
    </header>

    <div
      class="dropzone"
      :class="{
        'dropzone--active': dragActive,
        'dropzone--disabled': isBusy
      }"
      @dragenter.prevent="onDragEnter"
      @dragover.prevent="onDragOver"
      @dragleave.prevent="onDragLeave"
      @drop.prevent="onDrop"
    >
      <input
        ref="fileInput"
        type="file"
        class="sr-only"
        accept=".xlsx,.xls"
        @change="onFileChange"
      />
      <div class="dropzone__content">
        <svg
          class="dropzone__icon"
          viewBox="0 0 24 24"
          role="presentation"
          aria-hidden="true"
        >
          <path
            d="M12 16l-4-4h3V4h2v8h3l-4 4zm-7 2v2h14v-2H5z"
            fill="currentColor"
          />
        </svg>
        <p class="dropzone__title">拖拽 Excel 文件到此处</p>
        <p class="dropzone__subtitle">
          或
          <button class="link-button" type="button" @click="triggerFileDialog">
            点击选择
          </button>
          单个文件
        </p>
        <p class="dropzone__hint">{{ dragHint }}</p>
        <p class="dropzone__meta">允许类型：{{ ACCEPTED_EXTENSIONS.join(' / ') }}</p>
      </div>
    </div>

    <div class="status-panel" v-if="status === 'loading'">
      <p>�-؄��-��o"���z?��O�?s�,,�?��o?���'a?,</p>
    </div>
    <div class="status-panel status-panel--error" v-if="errorMessage">
      <p>{{ errorMessage }}</p>
    </div>

  </section>
</template>
