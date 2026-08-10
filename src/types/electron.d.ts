/**
 * Electron 桌面桥类型声明
 * ------------------------------------------------------------
 * 渲染进程通过 window.inkflow（由 electron/preload.cjs 的 contextBridge
 * 注入）访问原生文件对话框与菜单动作。本文件仅为类型，不包含实现。
 *
 * 在纯浏览器环境（npm run dev / 现有 e2e 测试跑在 Playwright 时）
 * window.inkflow 不存在，业务代码需用 inkflow.isElectron 或
 * typeof window !== 'undefined' && window.inkflow 做存在性判断并回落到
 * 浏览器行为（拖拽上传 + 浏览器下载）。
 */

interface InkflowOpenResult {
  name: string;
  filePath: string;
  data: Uint8Array;
  size: number;
}

interface InkflowSaveResult {
  filePath: string;
  name: string;
}

interface InkflowSavePayload {
  defaultName?: string;
  data: Uint8Array;
  saveAs?: boolean;
}

type InkflowMenuAction = 'new' | 'open' | 'save' | 'saveAs';

/** electron/preload.cjs 暴露的 window.inkflow 完整形状 */
interface WindowInkflow {
  isElectron: true;

  /** 弹出系统「打开」对话框并读取 .docx；取消返回 null，出错返回 { error } */
  openDocument(): Promise<InkflowOpenResult | null | { error: string }>;

  /** 保存/另存为 .docx；取消返回 null */
  saveDocument(payload: InkflowSavePayload): Promise<InkflowSaveResult | null>;

  /** 记录当前打开文档的文件路径，供「保存」直接覆盖；新建/另存为后调用以重设 */
  setCurrentFile(filePath: string | null): Promise<unknown>;

  /** 订阅原生菜单/快捷键动作，返回取消订阅函数 */
  onMenu(callback: (action: InkflowMenuAction) => void): () => void;

  /** 订阅从命令行/文件关联打开的 .docx，返回取消订阅函数 */
  onOpenResult(
    callback: (result: InkflowOpenResult | { error: string }) => void
  ): () => void;
}

interface Window {
  inkflow?: WindowInkflow;
}
