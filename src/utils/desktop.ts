/**
 * 桌面（Electron）能力统一入口
 * ------------------------------------------------------------
 * 把渲染进程对 window.inkflow（electron/preload.cjs）的调用集中在这里，
 * 并让所有方法在「纯浏览器」环境（npm run dev / Playwright e2e）下优雅回退。
 *
 * 关键思想：Electron 只是改变了「文件从哪来 / 输出到哪去」这两个 I/O 边界，
 * 中间的解析（parseDocx）与导出（jsonToDocxBlob）流水线完全复用原有实现——
 * 打开时把字节包装成 File 交给 parseDocx，保存时把 docx.js 产出的 Blob 字节
 * 经 IPC 写盘。
 */
import { parseDocx } from '../parser/parseDocx';
import { jsonToDocxBlob } from '../export/exportDocx';
import type { DocxComment } from '../parser/ooxml';

function api(): WindowInkflow | undefined {
  return typeof window !== 'undefined' ? window.inkflow : undefined;
}

/** IPC 跨进程返回的 Uint8Array 可能是 ArrayBufferLike（SharedArrayBuffer），
 *  这里规整成 ArrayBuffer 背书的视图，才能安全地作为 Blob/File 的 BlobPart */
function toArrayBufferView(data: Uint8Array): Uint8Array<ArrayBuffer> {
  return Uint8Array.from(data);
}

/** 是否运行在 Electron 桌面环境 */
export const isElectron = (): boolean => !!api();

export interface OpenedDocx {
  json: any;
  fileName: string;
  warnings: string[];
  comments: DocxComment[];
  pageSetup: any;
  header: any | null;
  footer: any | null;
  sections: any[];
}

/**
 * 通过系统「打开」对话框选择一个 .docx 并完成解析。
 * 非 Electron 环境返回 null（由调用方走浏览器拖拽/选择上传）。
 * 失败时抛出错误。
 */
export async function openDocxViaDialog(): Promise<OpenedDocx | null> {
  const bridge = api();
  if (!bridge) return null;

  const result = await bridge.openDocument();
  if (!result) return null;
  if ('error' in result) throw new Error(result.error);

  // 字节打包成 File → 复用原有 parseDocx 流水线（含 Worker 化解析）
  const file = new File([toArrayBufferView(result.data)], result.name, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await bridge.setCurrentFile(result.filePath);
  const { json, warnings, comments, pageSetup, header, footer, sections } =
    await parseDocx(file);
  return {
    json,
    fileName: result.name,
    warnings,
    comments,
    pageSetup,
    header,
    footer,
    sections,
  };
}

/**
 * 通过系统「保存 / 另存为」对话框把当前文档导出为 .docx 并写盘。
 * 返回实际保存的文件名（取消返回 null）；非 Electron 环境返回 null。
 */
export async function saveDocxViaDialog(
  docJson: any,
  fileName: string,
  comments: DocxComment[],
  saveAs = false
): Promise<string | null> {
  const bridge = api();
  if (!bridge) return null;

  const blob = await jsonToDocxBlob(docJson, comments);
  const buf = await blob.arrayBuffer();
  const result = await bridge.saveDocument({
    defaultName: fileName,
    data: new Uint8Array(buf),
    saveAs,
  });
  return result?.name ?? null;
}

/** 新建空白文档时清空「当前文件」记录，使后续「保存」走另存为 */
export async function resetCurrentFile(): Promise<void> {
  const bridge = api();
  if (bridge) await bridge.setCurrentFile(null);
}

/** 订阅原生菜单/快捷键动作（new|open|save|saveAs），返回取消订阅函数 */
export function onMenu(
  cb: (action: 'new' | 'open' | 'save' | 'saveAs') => void
): () => void {
  const bridge = api();
  if (!bridge) return () => {};
  return bridge.onMenu(cb);
}

/** 订阅从命令行/文件关联打开的 .docx */
export function onOpenResult(
  cb: (result: InkflowOpenResult | { error: string }) => void
): () => void {
  const bridge = api();
  if (!bridge) return () => {};
  return bridge.onOpenResult(cb);
}
