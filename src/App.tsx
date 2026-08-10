import { createSignal, Show, onMount } from 'solid-js';
import { UploadPanel } from './upload/UploadPanel';
import { EditorPane } from './editor/EditorPane';
import { PreviewPane } from './preview/PreviewPane';
import { exportAndDownloadDocx } from './export/exportDocx';
import { applyAIPatch } from './utils/applyAIPatch';
import { EMPTY_DOC } from './schema';
import {
  isElectron,
  onMenu,
  onOpenResult,
  openDocxViaDialog,
  saveDocxViaDialog,
  resetCurrentFile,
} from './utils/desktop';
import type { DocxComment } from './parser/ooxml';

type Stage = 'upload' | 'edit' | 'preview';

function App() {
  const [docJson, setDocJson] = createSignal<any | null>(null);
  const [fileName, setFileName] = createSignal('未命名文档.docx');
  const [stage, setStage] = createSignal<Stage>('upload');
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

  const handleExport = async (saveAs = false) => {
    const current = docJson();
    if (!current) return;
    setBusy('export');
    try {
      if (isElectron()) {
        // 桌面环境：通过系统「保存/另存为」对话框写盘
        const savedName = await saveDocxViaDialog(
          current,
          fileName(),
          comments(),
          saveAs
        );
        if (savedName) {
          setFileName(savedName);
          setPulseExport(true);
          showToast('已保存 DOCX');
          setTimeout(() => setPulseExport(false), 1600);
        }
      } else {
        // 浏览器环境：沿用原有 file-saver 下载
        await exportAndDownloadDocx(current, fileName(), comments());
        setPulseExport(true);
        showToast('已导出 DOCX');
        setTimeout(() => setPulseExport(false), 1600);
      }
    } catch (err) {
      showToast(
        `导出失败：${err instanceof Error ? err.message : String(err)}`,
        3000
      );
    } finally {
      setBusy('');
    }
  };

  // 新建空白文档
  const handleNew = async () => {
    if (isElectron()) await resetCurrentFile();
    setDocJson(EMPTY_DOC);
    setFileName('未命名文档.docx');
    setWarnings([]);
    setComments([]);
    setLoadKey((k) => k + 1);
    setStage('edit');
    showToast('已新建空白文档');
  };

  // 用解析结果进入编辑态（供桌面打开/命令行打开复用）
  const applyParsed = (
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

  // 桌面集成：订阅原生菜单快捷键 + 命令行/文件关联打开的 .docx
  onMount(() => {
    const offMenu = onMenu((action) => {
      if (action === 'new') void handleNew();
      else if (action === 'open') void openViaMenu();
      else if (action === 'save') void handleExport(false);
      else if (action === 'saveAs') void handleExport(true);
    });

    const offOpen = onOpenResult(async (result) => {
      if (!result || 'error' in result) return;
      // 规整成 ArrayBuffer 背书的视图后才能作为 File 的 BlobPart
      const view: Uint8Array<ArrayBuffer> = Uint8Array.from(result.data);
      const file = new File([view], result.name, {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      try {
        const { parseDocx } = await import('./parser/parseDocx');
        const { json, warnings, comments } = await parseDocx(file);
        applyParsed(json, result.name, warnings, comments);
      } catch (err) {
        showToast(
          `打开失败：${err instanceof Error ? err.message : String(err)}`,
          3000
        );
      }
    });

    return () => {
      offMenu();
      offOpen();
    };
  });

  // 供菜单「打开」复用：与 UploadPanel 共用同一套桌面解析入口
  const openViaMenu = async () => {
    try {
      const opened = await openDocxViaDialog();
      if (opened) applyParsed(opened.json, opened.fileName, opened.warnings, opened.comments);
    } catch (err) {
      showToast(
        `打开失败：${err instanceof Error ? err.message : String(err)}`,
        3000
      );
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

  const activeDot = 'bg-accent shadow-[0_0_0_3px_var(--color-accent-wash)]';
  const doneDot = 'bg-accent-soft';
  const inactiveDot = 'bg-line-strong';

  const pipelineStep = 'flex items-center gap-1.5 px-2 py-1 rounded-full text-xs font-mono transition-colors';

  return (
    <div class="h-screen flex flex-col bg-canvas">
      <header class="flex items-center gap-7 px-6 py-3.5 bg-paper border-b border-line">
        <div class="flex items-center gap-2 flex-shrink-0">
          <span class="font-display font-semibold text-xl tracking-tight text-ink-1">
            InkFlow
          </span>
        </div>

        <nav class="flex items-center gap-2 flex-1 justify-center text-ink-3" aria-label="处理流程">
          <span
            class={`${pipelineStep} ${hasDoc() ? 'text-ink-2' : stage() === 'upload' ? 'text-accent-ink font-medium' : ''}`}
          >
            <span class={`w-1.5 h-1.5 rounded-full transition-all ${hasDoc() ? doneDot : stage() === 'upload' ? activeDot : inactiveDot}`} />
            上传 DOCX
          </span>
          <span class="w-7 h-px bg-[repeating-linear-gradient(to_right,var(--color-line-strong)_0,var(--color-line-strong)_3px,transparent_3px,transparent_6px)]" />
          <span
            class={`${pipelineStep} ${stage() !== 'upload' ? 'text-accent-ink font-medium' : ''}`}
          >
            <span class={`w-1.5 h-1.5 rounded-full transition-all ${stage() !== 'upload' ? activeDot : inactiveDot}`} />
            结构化 JSON
          </span>
          <span class="w-7 h-px bg-[repeating-linear-gradient(to_right,var(--color-line-strong)_0,var(--color-line-strong)_3px,transparent_3px,transparent_6px)]" />
          <span
            class={`${pipelineStep} ${pulseExport() ? 'text-accent-ink font-medium' : ''}`}
          >
            <span class={`w-1.5 h-1.5 rounded-full transition-all ${pulseExport() ? `${activeDot} animate-[pipeline-pulse_0.8s_ease-in-out_2]` : inactiveDot}`} />
            导出 DOCX
          </span>
        </nav>

        <div class="flex items-center gap-2.5 flex-shrink-0">
          <div class="flex bg-surface-1 rounded-lg p-0.5 mr-1.5" role="tablist">
            <button
              role="tab"
              aria-selected={stage() === 'upload'}
              class={`px-3.5 py-1.5 rounded-md text-[13px] font-semibold transition-all ${stage() === 'upload' ? 'bg-paper text-ink-1 shadow-sm' : 'text-ink-2 hover:text-ink-1'}`}
              onClick={() => setStage('upload')}
            >
              上传
            </button>
            <button
              role="tab"
              aria-selected={stage() === 'edit'}
              class={`px-3.5 py-1.5 rounded-md text-[13px] font-semibold transition-all disabled:opacity-45 disabled:cursor-not-allowed ${stage() === 'edit' ? 'bg-paper text-ink-1 shadow-sm' : 'text-ink-2 hover:text-ink-1'}`}
              disabled={!hasDoc()}
              onClick={() => setStage('edit')}
            >
              编辑
            </button>
            <button
              role="tab"
              aria-selected={stage() === 'preview'}
              class={`px-3.5 py-1.5 rounded-md text-[13px] font-semibold transition-all disabled:opacity-45 disabled:cursor-not-allowed ${stage() === 'preview' ? 'bg-paper text-ink-1 shadow-sm' : 'text-ink-2 hover:text-ink-1'}`}
              disabled={!hasDoc()}
              onClick={() => setStage('preview')}
            >
              预览
            </button>
          </div>
          <Show when={isElectron()}>
            <button
              type="button"
              class="px-3.5 py-2 rounded-lg text-[13px] font-semibold border border-line-strong bg-paper text-ink-1 transition-all hover:border-accent-soft hover:bg-accent-wash disabled:opacity-45 disabled:cursor-not-allowed"
              disabled={busy() !== ''}
              onClick={handleNew}
              title="新建空白文档 (Ctrl+N)"
            >
              新建
            </button>
          </Show>
          <button
            type="button"
            class="px-3.5 py-2 rounded-lg text-[13px] font-semibold border border-line-strong bg-paper text-ink-1 transition-all hover:border-accent-soft hover:bg-accent-wash disabled:opacity-45 disabled:cursor-not-allowed"
            disabled={!hasDoc() || busy() !== ''}
            onClick={handleAIOutline}
            title="AI 扩展接口示例：根据标题自动生成大纲（applyAIPatch）"
          >
            {busy() === 'ai' ? '生成中…' : '生成大纲'}
          </button>
          <button
            type="button"
            class="px-3.5 py-2 rounded-lg text-[13px] font-semibold border border-accent bg-accent text-white transition-all hover:bg-accent-ink hover:border-accent-ink disabled:opacity-45 disabled:cursor-not-allowed"
            disabled={!hasDoc() || busy() !== ''}
            onClick={() => handleExport(false)}
          >
            {busy() === 'export' ? '导出中…' : '导出 DOCX'}
          </button>
        </div>
      </header>

      <Show when={warnings().length > 0 && stage() !== 'upload'}>
        <div class="bg-warning/15 text-warning-ink text-xs px-6 py-2 border-b border-warning/40 dark:bg-warning/10 dark:text-warning-ink dark:border-warning/40">
          还原提示（结构优先，样式可能与原文件存在差异）：
          {warnings().slice(0, 2).join('；')}
          {warnings().length > 2 ? ` 等 ${warnings().length} 项` : ''}
        </div>
      </Show>

      <main class="flex-1 min-h-0 flex flex-col">
        <Show when={stage() === 'upload'}>
          <UploadPanel onParsed={handleParsed} />
        </Show>
        <Show when={stage() === 'edit' && hasDoc()}>
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
          <PreviewPane docJson={docJson()} initialComments={comments()} onCommentsChange={setComments} />
        </Show>
      </main>

      <Show when={toast()}>
        <div class="fixed bottom-6 left-1/2 -translate-x-1/2 bg-ink-1 text-white px-5 py-2.5 rounded-lg text-sm shadow-xl animate-[toast-in_0.2s_ease]">
          {toast()}
        </div>
      </Show>
    </div>
  );
}

export default App;
