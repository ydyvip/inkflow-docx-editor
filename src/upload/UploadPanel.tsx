import { createSignal } from 'solid-js';
import { parseDocx } from '../parser/parseDocx';
import { EMPTY_DOC } from '../schema';
import type { DocxComment } from '../parser/ooxml';
import './upload.css';

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
    <div class="upload-panel">
      <div
        class={`dropzone${isDragging() ? ' is-dragging' : ''}${status() === 'parsing' ? ' is-busy' : ''}`}
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
        <div class="dropzone-icon" aria-hidden>
          {status() === 'parsing' ? '⟳' : '⇪'}
        </div>
        <div class="dropzone-title">
          {status() === 'parsing'
            ? '正在解析 DOCX…'
            : '拖拽 .docx 文件到此处，或点击选择'}
        </div>
        <div class="dropzone-sub">
          DOCX → mammoth → HTML(中间态) → DOM → ProseMirror JSON
        </div>
        {status() === 'error' && (
          <div class="dropzone-error">解析失败：{error()}</div>
        )}
      </div>

      <button type="button" class="ghost-btn" onClick={startBlank}>
        或从空白文档开始 →
      </button>
    </div>
  );
}
