import { createSignal, Show } from 'solid-js';
import { UploadPanel } from './upload/UploadPanel';
import { EditorPane } from './editor/EditorPane';
import { PreviewPane } from './preview/PreviewPane';
import { exportAndDownloadDocx } from './export/exportDocx';
import { applyAIPatch } from './utils/applyAIPatch';
import type { DocxComment } from './parser/ooxml';
import './App.css';

type Stage = 'upload' | 'edit' | 'preview';

function App() {
  const [docJson, setDocJson] = createSignal<any | null>(null);
  const [fileName, setFileName] = createSignal('未命名文档.docx');
  const [stage, setStage] = createSignal<Stage>('upload');
  // loadKey 从 1 开始（非 0），配合 <Show keyed> 在每次加载新文档时强制重新挂载 EditorPane
  const [loadKey, setLoadKey] = createSignal(1);
  const [warnings, setWarnings] = createSignal<string[]>([]);
  const [comments, setComments] = createSignal<DocxComment[]>([]);
  const [busy, setBusy] = createSignal<'' | 'export' | 'ai'>('');
  const [pulseExport, setPulseExport] = createSignal(false);
  const [toast, setToast] = createSignal<string | null>(null);

  const hasDoc = () => docJson() !== null;

  const showToast = (msg: string, ms = 2200) => {
    setToast(msg);
    setTimeout(() => setToast(null), ms);
  };

  const handleParsed = (
    json: any,
    name: string,
    msgs: string[],
    parsedComments: DocxComment[]
  ) => {
    setDocJson(json);
    setFileName(name);
    setWarnings(msgs);
    setComments(parsedComments);
    setLoadKey((k) => k + 1);
    setStage('edit');
  };

  const handleChange = (json: any) => {
    setDocJson(json);
  };

  const handleExport = async () => {
    const current = docJson();
    if (!current) return;
    setBusy('export');
    try {
      await exportAndDownloadDocx(current, fileName(), comments());
      setPulseExport(true);
      showToast('已导出 DOCX');
      setTimeout(() => setPulseExport(false), 1600);
    } catch (err) {
      showToast(
        `导出失败：${err instanceof Error ? err.message : String(err)}`,
        3000
      );
    } finally {
      setBusy('');
    }
  };

  const handleAIOutline = async () => {
    const current = docJson();
    if (!current) return;
    setBusy('ai');
    try {
      const patched = await applyAIPatch(current, '生成大纲');
      setDocJson(patched);
      setLoadKey((k) => k + 1);
      setStage('edit');
      showToast('已插入自动大纲');
    } finally {
      setBusy('');
    }
  };

  return (
    <div class="app-shell">
      <header class="app-header">
        <div class="brand">
          <span class="brand-mark" aria-hidden>
            🖋
          </span>
          <span class="brand-name">InkFlow</span>
        </div>

        <nav class="pipeline" aria-label="处理流程">
          <span
            class={`pipeline-step${hasDoc() ? ' is-done' : stage() === 'upload' ? ' is-active' : ''}`}
          >
            <span class="pipeline-dot" />
            上传 DOCX
          </span>
          <span class="pipeline-line" />
          <span
            class={`pipeline-step${stage() !== 'upload' ? ' is-active' : ''}`}
          >
            <span class="pipeline-dot" />
            结构化 JSON
          </span>
          <span class="pipeline-line" />
          <span
            class={`pipeline-step${pulseExport() ? ' is-active is-pulse' : ''}`}
          >
            <span class="pipeline-dot" />
            导出 DOCX
          </span>
        </nav>

        <div class="header-actions">
          <div class="tabs" role="tablist">
            <button
              role="tab"
              aria-selected={stage() === 'upload'}
              class={`tab-btn${stage() === 'upload' ? ' is-active' : ''}`}
              onClick={() => setStage('upload')}
            >
              上传
            </button>
            <button
              role="tab"
              aria-selected={stage() === 'edit'}
              class={`tab-btn${stage() === 'edit' ? ' is-active' : ''}`}
              disabled={!hasDoc()}
              onClick={() => setStage('edit')}
            >
              编辑
            </button>
            <button
              role="tab"
              aria-selected={stage() === 'preview'}
              class={`tab-btn${stage() === 'preview' ? ' is-active' : ''}`}
              disabled={!hasDoc()}
              onClick={() => setStage('preview')}
            >
              预览
            </button>
          </div>
          <button
            type="button"
            class="ai-btn"
            disabled={!hasDoc() || busy() !== ''}
            onClick={handleAIOutline}
            title="AI 扩展接口示例：根据标题自动生成大纲（applyAIPatch）"
          >
            {busy() === 'ai' ? '生成中…' : '✨ 生成大纲'}
          </button>
          <button
            type="button"
            class="export-btn"
            disabled={!hasDoc() || busy() !== ''}
            onClick={handleExport}
          >
            {busy() === 'export' ? '导出中…' : '导出 DOCX'}
          </button>
        </div>
      </header>

      <Show when={warnings().length > 0 && stage() !== 'upload'}>
        <div class="warning-bar">
          还原提示（结构优先，样式可能与原文件存在差异）：
          {warnings().slice(0, 2).join('；')}
          {warnings().length > 2 ? ` 等 ${warnings().length} 项` : ''}
        </div>
      </Show>

      <main class="app-main">
        <Show when={stage() === 'upload'}>
          <UploadPanel onParsed={handleParsed} />
        </Show>
        <Show when={stage() === 'edit' && hasDoc()}>
          {/* keyed Show：每次加载新文档（上传 / AI 大纲）时强制重新挂载编辑器，
              对应原方案里"用最新 JSON 重建 EditorState"的语义 */}
          <Show when={loadKey()} keyed>
            {(_key) => (
              <EditorPane
                initialDoc={docJson()}
                initialComments={comments()}
                onChange={handleChange}
                onCommentsChange={setComments}
              />
            )}
          </Show>
        </Show>
        <Show when={stage() === 'preview' && hasDoc()}>
          <PreviewPane docJson={docJson()} />
        </Show>
      </main>

      <Show when={toast()}>
        <div class="toast">{toast()}</div>
      </Show>
    </div>
  );
}

export default App;
