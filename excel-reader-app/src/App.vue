
<template>
  <div class="app-shell">
    <header class="hero">
      <div class="hero__content">
        <p class="eyebrow">Vue 3 · Vite · SheetJS</p>
        <h1>Excel 在线阅读器</h1>
        <p class="hero__lead">
          快速上传并解析 Excel，所有数据都在浏览器端即时处理，适用于演示、校验与临时数据核对场景。
        </p>
        <ul class="hero__steps" aria-label="使用步骤">
          <li>上传 .xlsx / .xls 文件</li>
          <li>浏览或检索工作表数据</li>
          <li>导出 JSON / CSV 或重置</li>
        </ul>
      </div>
    </header>

    <main class="content-area">
      <ExcelReader @result-generated="handleResultText" />

      <section class="guide-card" aria-labelledby="guide-title">
        <header class="guide-card__header">
          <div>
            <p class="eyebrow">数据摘要</p>
            <h2 id="guide-title">转换结果</h2>
            <p>解析成功后会在此生成一份标准文案，可复制给其他同事或粘贴进工单。</p>
          </div>
          <button
            type="button"
            class="copy-btn"
            :class="{ 'copy-btn--success': copySuccess }"
            :disabled="!resultText"
            @click="copyResult"
          >
            {{ copySuccess ? '已复制' : '复制' }}
          </button>
        </header>
        <textarea
          :value="resultText"
          class="guide-card__editor"
          rows="16"
          aria-label="转换结果文本"
          placeholder="解析 Excel 后将在此生成标准文本，方便复制分享。"
          readonly
        ></textarea>
      </section>
    </main>

    <footer class="app-footer">
      <p>纯前端实现 · 本地缓存最近一次解析结果，刷新后仍可继续查看。</p>
    </footer>
  </div>
</template>

<script setup>
import { ref } from 'vue';
import ExcelReader from './components/ExcelReader.vue';

const resultText = ref('');
const copySuccess = ref(false);
let copyTimer;

function handleResultText(text) {
  resultText.value = text ?? '';
  copySuccess.value = false;
  clearTimeout(copyTimer);
}

async function copyResult() {
  if (!resultText.value) return;
  try {
    await navigator.clipboard.writeText(resultText.value);
    copySuccess.value = true;
  } catch (error) {
    copySuccess.value = false;
    console.error('复制失败', error);
  } finally {
    clearTimeout(copyTimer);
    copyTimer = setTimeout(() => {
      copySuccess.value = false;
    }, 2000);
  }
}
</script>
