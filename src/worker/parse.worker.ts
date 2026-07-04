/// <reference lib="webworker" />
/**
 * Parser Worker
 * ------------------------------------------------------------
 * DOCX 解析里最重的部分（zip 解包 + XML 遍历）在这里完成，避免大文件
 * 解析阻塞主线程动画帧（§9.1 / §9.2 Worker解析）。
 *
 * DOMParser 在现代浏览器的 Worker 环境下可用，因此这次连"生成 JSON"
 * 这一步也一起放进了 Worker——不再需要像 mammoth 方案那样，回到主线程
 * 用浏览器 DOM 把 HTML 转成 ProseMirror JSON。
 */
import { parseDocxFile } from "../parser/ooxml";

export interface ParseWorkerRequest {
  id: string;
  arrayBuffer: ArrayBuffer;
}

export interface ParseWorkerResponse {
  id: string;
  ok: boolean;
  json?: any;
  comments?: { id: number; author: string; date: string | null; text: string }[];
  warnings?: string[];
  error?: string;
}

self.onmessage = async (event: MessageEvent<ParseWorkerRequest>) => {
  const { id, arrayBuffer } = event.data;
  try {
    const result = await parseDocxFile(arrayBuffer);
    const response: ParseWorkerResponse = {
      id,
      ok: true,
      json: result.json,
      comments: result.comments,
      warnings: result.warnings,
    };
    (self as unknown as Worker).postMessage(response);
  } catch (err) {
    const response: ParseWorkerResponse = {
      id,
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    };
    (self as unknown as Worker).postMessage(response);
  }
};

export {};
