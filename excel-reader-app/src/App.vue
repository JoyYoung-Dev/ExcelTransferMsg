
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
      <ExcelReader />

      <section class="guide-card" aria-labelledby="guide-title">
        <div class="guide-card__header">
          <p class="eyebrow">流程提醒</p>
          <h2 id="guide-title">操作指南</h2>
          <p>依照以下步骤即可在几秒钟内完成 Excel 在线查看与导出。</p>
        </div>
        <ol class="guide-list">
          <li><strong>准备文件。</strong> 支持 Office 及 WPS 导出的 .xlsx / .xls 单文件。</li>
          <li><strong>拖拽或点击上传。</strong> 文件体积建议控制在 5 MB 以内以获得最佳体验。</li>
          <li><strong>等待解析。</strong> SheetJS 在浏览器内完成解码，不会上传到服务器。</li>
          <li><strong>筛选与导出。</strong> 可搜索关键字、切换工作表，并导出 JSON 或 CSV。</li>
          <li><strong>重置。</strong> 需要重新开始时点击重置即可清理缓存及本地存储。</li>
        </ol>
      </section>
    </main>

    <footer class="app-footer">
      <p>纯前端实现 · 本地缓存最近一次解析结果，刷新后仍可继续查看。</p>
    </footer>
  </div>
</template>

<script setup>
import ExcelReader from './components/ExcelReader.vue';
</script>
