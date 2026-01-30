
<script setup>
import { computed, onMounted, ref, watch } from 'vue';
import * as XLSX from 'xlsx';

const ACCEPTED_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel'
];
const ACCEPTED_EXTENSIONS = ['.xlsx', '.xls'];
const STORAGE_KEY = 'excel-reader-cache';
const MAX_CACHE_ROWS_PER_SHEET = 500;
const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB is a safe browser limit

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
const canDownload = computed(() => hasData.value && !isBusy.value);

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
        rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], {
          header: 1,
          blankrows: false,
          defval: ''
        })
      }));

      if (!workbookSheets.length) {
        throw new Error('EMPTY_WORKBOOK');
      }

      sheets.value = workbookSheets;
      selectedSheet.value = workbookSheets[0].name;
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
}

function downloadJSON() {
  if (!canDownload.value) return;
  const payload = {
    generatedAt: new Date().toISOString(),
    file: fileInfo.value,
    sheet: currentSheet.value?.name ?? '',
    rows: currentRows.value
  };
  triggerDownload(
    `${(fileInfo.value?.name ?? 'excel-data').replace(/\.[^.]+$/, '')}.json`,
    new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' })
  );
}

function downloadCSV() {
  if (!canDownload.value || !currentRows.value.length) return;
  const worksheet = XLSX.utils.aoa_to_sheet(currentRows.value);
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  triggerDownload(
    `${currentSheet.value?.name ?? 'worksheet'}.csv`,
    new Blob([csv], { type: 'text/csv;charset=utf-8;' })
  );
}

function triggerDownload(filename, blob) {
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);
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
      <p>文件正在解析，通常只需数秒。</p>
    </div>
    <div class="status-panel status-panel--error" v-if="errorMessage">
      <p>{{ errorMessage }}</p>
    </div>

    <section class="file-summary" v-if="fileInfo">
      <div class="file-summary__grid">
        <article>
          <span class="label">文件名</span>
          <p>{{ fileInfo.name }}</p>
        </article>
        <article>
          <span class="label">大小</span>
          <p>{{ fileInfo.sizeLabel }}</p>
        </article>
        <article>
          <span class="label">工作表</span>
          <p>{{ fileInfo.sheetCount }} 个</p>
        </article>
        <article>
          <span class="label">最近修改</span>
          <p>{{ fileInfo.lastModifiedLabel }}</p>
        </article>
      </div>
      <div class="file-summary__meta">
        <p v-if="lastUpdatedText">解析时间：{{ lastUpdatedText }}</p>
        <button class="btn btn--ghost" type="button" @click="resetAll">
          重置
        </button>
      </div>
    </section>

    <section v-if="hasData" class="sheet-tools" aria-label="数据工具栏">
      <label class="field">
        <span>工作表</span>
        <select v-model="selectedSheet">
          <option
            v-for="option in sheetOptions"
            :key="option.value"
            :value="option.value"
          >
            {{ option.label }}
          </option>
        </select>
      </label>

      <label class="field">
        <span>关键字筛选</span>
        <input
          type="search"
          v-model.trim="searchTerm"
          placeholder="输入任意文本以过滤"
        />
      </label>

      <div class="sheet-tools__actions">
        <button
          class="btn btn--primary"
          type="button"
          @click="downloadCSV"
          :disabled="!canDownload"
        >
          导出 CSV
        </button>
        <button
          class="btn btn--secondary"
          type="button"
          @click="downloadJSON"
          :disabled="!canDownload"
        >
          导出 JSON
        </button>
        <button class="btn btn--ghost" type="button" @click="resetAll">
          清除
        </button>
      </div>
    </section>

    <div v-if="hasData" class="table-wrapper">
      <div class="table-meta">
        <p>共 {{ summary.totalRows }} 行 / {{ summary.columns }} 列</p>
        <p v-if="summary.filteredRows !== summary.totalRows">
          已筛选出 {{ summary.filteredRows }} 行
        </p>
      </div>
      <div class="table-scroll" role="region" aria-live="polite">
        <table>
          <thead>
            <tr>
              <th v-for="(header, index) in normalizedHeaders" :key="`header-${index}`">
                {{ header }}
              </th>
            </tr>
          </thead>
          <tbody>
            <tr v-if="!filteredRows.length">
              <td :colspan="Math.max(normalizedHeaders.length, 1)">
                没有匹配的数据，请更换关键字或重置筛选。
              </td>
            </tr>
            <tr v-for="(row, rowIndex) in filteredRows" :key="`row-${rowIndex}`">
              <td
                v-for="(cell, cellIndex) in row"
                :key="`cell-${rowIndex}-${cellIndex}`"
              >
                {{ formatCell(cell) }}
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <p v-else class="placeholder-text">
      上传 Excel 文件后，这里将呈现解析结果并支持导出与筛选。
    </p>
  </section>
</template>
