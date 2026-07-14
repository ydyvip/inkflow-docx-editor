import { createSignal } from 'solid-js';
import { parseDocx } from '../parser/parseDocx';
import { EMPTY_DOC } from '../schema';
import type { DocxComment } from '../parser/ooxml';

interface UploadPanelProps {
  onParsed: (
    json: any,
    fileName: string,
    warnings: string[],
    comments: DocxComment[]
  ) => void;
}

type Status = 'idle' | 'parsing' | 'error';

export function UploadPanel(props: UploadPanelProps) {
  const [status, setStatus] = createSignal<Status>('idle');
  const [error, setError] = createSignal<string | null>(null);
  const [isDragging, setDragging] = createSignal(false);
  let inputEl: HTMLInputElement | undefined;

  const handleFile = async (file: File | undefined) => {
    if (!file) return;
    setStatus('parsing');
    setError(null);
    try {
      const { json, warnings, comments } = await parseDocx(file);
      props.onParsed(json, file.name, warnings, comments);
      setStatus('idle');
    } catch (err) {
      setStatus('error');
      setError(err instanceof Error ? err.message : '解析失败');
    }
  };

  const startBlank = () => props.onParsed(EMPTY_DOC, '未命名文档.docx', [], []);

  return (
    <div class="h-full flex flex-col items-center justify-center gap-5 p-10 bg-canvas">
      <div
        class={`w-full max-w-[560px] border-2 border-dashed rounded-xl bg-paper p-14 text-center cursor-pointer transition-all hover:border-accent hover:bg-accent-wash focus-visible:border-accent focus-visible:bg-accent-wash focus-visible:outline-none ${isDragging() ? 'border-accent bg-accent-wash scale-[1.01]' : 'border-line-strong'} ${status() === 'parsing' ? 'opacity-85 cursor-progress' : ''}`}
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragging(false);
          handleFile(e.dataTransfer?.files?.[0]);
        }}
        onClick={() => inputEl?.click()}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') inputEl?.click();
        }}
      >
        <input
          ref={inputEl}
          type="file"
          accept=".docx"
          hidden
          onChange={(e) => handleFile(e.currentTarget.files?.[0])}
        />
        <div class={`text-4xl text-accent-ink mb-3.5 ${status() === 'parsing' ? 'animate-spin inline-block' : ''}`} aria-hidden>
          {status() === 'parsing' ? '⟳' : '⇪'}
        </div>
        <div class="font-display font-semibold text-lg text-ink-1 mb-1.5">
          {status() === 'parsing'
            ? '正在解析 DOCX…'
            : '拖拽 .docx 文件到此处，或点击选择'}
        </div>
        <div class="font-mono text-[11.5px] text-ink-3 tracking-wide">
          DOCX → mammoth → HTML(中间态) → DOM → ProseMirror JSON
        </div>
        {status() === 'error' && (
          <div class="mt-4 text-danger text-sm font-semibold">解析失败：{error()}</div>
        )}
      </div>

      <button type="button" class="bg-transparent border-0 text-ink-2 text-[13.5px] cursor-pointer underline underline-offset-3 decoration-line-strong hover:text-accent-ink" onClick={startBlank}>
        或从空白文档开始
      </button>
    </div>
  );
}
