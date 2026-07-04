/**
 * Parser 模块（核心）
 * ------------------------------------------------------------
 * 流程：DOCX → Worker（zip 解包 + OOXML 遍历）→ ProseMirror JSON + 批注列表
 *
 * 约束：
 *   - 大文件必须走 Web Worker（§9.1）
 *   - 对外只暴露 JSON（系统唯一真相源）+ 批注元数据，不暴露中间态
 */
import type { ParseWorkerRequest, ParseWorkerResponse } from "../worker/parse.worker";
import type { DocxComment } from "./ooxml";

export interface ParseResult {
  json: any; // ProseMirror JSON Document（§5.1）—— 系统唯一真相源
  comments: DocxComment[];
  warnings: string[];
}

const MAX_MAIN_THREAD_BYTES = 2 * 1024 * 1024; // 2MB 以下允许主线程兜底

function convertViaWorker(arrayBuffer: ArrayBuffer): Promise<ParseResult> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("../worker/parse.worker.ts", import.meta.url), { type: "module" });
    const id = crypto.randomUUID();

    const timeout = setTimeout(() => {
      worker.terminate();
      reject(new Error("Worker 解析超时"));
    }, 30_000);

    worker.onmessage = (event: MessageEvent<ParseWorkerResponse>) => {
      if (event.data.id !== id) return;
      clearTimeout(timeout);
      worker.terminate();
      if (event.data.ok && event.data.json) {
        resolve({ json: event.data.json, comments: event.data.comments ?? [], warnings: event.data.warnings ?? [] });
      } else {
        reject(new Error(event.data.error ?? "Worker 解析失败"));
      }
    };
    worker.onerror = (e) => {
      clearTimeout(timeout);
      worker.terminate();
      reject(e.error ?? new Error("Worker 异常"));
    };

    const request: ParseWorkerRequest = { id, arrayBuffer };
    worker.postMessage(request, [arrayBuffer]);
  });
}

async function convertViaMainThread(arrayBuffer: ArrayBuffer): Promise<ParseResult> {
  const { parseDocxFile } = await import("./ooxml");
  const result = await parseDocxFile(arrayBuffer);
  return { json: result.json, comments: result.comments, warnings: result.warnings };
}

/** 校验：必须是 .docx（§6.1 Upload Module 规则）*/
export function validateDocxFile(file: File): void {
  const isDocxExt = file.name.toLowerCase().endsWith(".docx");
  const isDocxMime =
    file.type === "" ||
    file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (!isDocxExt || !isDocxMime) {
    throw new Error("仅支持 .docx 文件");
  }
}

/** 解析入口：File → ArrayBuffer → (Worker) OOXML → ProseMirror JSON + 批注 */
export async function parseDocx(file: File): Promise<ParseResult> {
  validateDocxFile(file);
  const arrayBuffer = await file.arrayBuffer();

  const preferWorker = typeof Worker !== "undefined" && arrayBuffer.byteLength > 0;
  if (preferWorker) {
    try {
      return await convertViaWorker(arrayBuffer.slice(0));
    } catch (workerErr) {
      if (arrayBuffer.byteLength > MAX_MAIN_THREAD_BYTES) throw workerErr;
      return convertViaMainThread(arrayBuffer);
    }
  }
  return convertViaMainThread(arrayBuffer);
}
