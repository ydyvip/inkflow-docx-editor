/**
 * AI 扩展接口（§11，预留）
 * ------------------------------------------------------------
 * 约定：applyAIPatch(jsonDoc, instruction) → jsonDoc
 * 输入输出都是 ProseMirror JSON —— 与系统唯一真相源保持一致，
 * 这样任何 AI 能力（改写 / 总结 / 优化结构）都可以作为一次
 * "transaction 前置生成器"接入，而不需要绕过 Editor 模块直接
 * 操作 DOCX 或 HTML。
 *
 * 本文件默认提供一个不依赖网络的本地示例实现（"生成大纲"），
 * 真实项目应把 runAIPatch 替换为对你自己的 LLM 服务的调用：
 *
 *   export async function runAIPatch(jsonDoc, instruction) {
 *     const res = await fetch("/api/ai-patch", {
 *       method: "POST",
 *       body: JSON.stringify({ jsonDoc, instruction }),
 *     });
 *     return res.json(); // 必须返回合法的 ProseMirror JSON
 *   }
 */

export type AIPatchFn = (jsonDoc: any, instruction: string) => Promise<any>;

function collectHeadings(node: any, out: { level: number; text: string }[]) {
  if (node.type === "heading") {
    const text = (node.content ?? []).map((c: any) => c.text ?? "").join("");
    if (text.trim()) out.push({ level: node.attrs?.level ?? 1, text });
  }
  for (const child of node.content ?? []) {
    if (child.type !== "text") collectHeadings(child, out);
  }
}

/**
 * 示例实现：基于文档现有标题，在文档最前面插入一个"目录大纲"提示块。
 * 演示"优化文档结构"这一类 AI 能力如何以 JSON→JSON 的形式接入系统，
 * 不依赖外部网络请求。
 */
async function localOutlinePatch(jsonDoc: any): Promise<any> {
  const headings: { level: number; text: string }[] = [];
  collectHeadings(jsonDoc, headings);

  if (headings.length === 0) {
    return jsonDoc; // 没有标题，不生成大纲
  }

  const outlineParagraphs = headings.map((h) => ({
    type: "paragraph",
    content: [{ type: "text", text: `${"　".repeat(Math.max(0, h.level - 1))}· ${h.text}` }],
  }));

  const outlineBlock = {
    type: "callout",
    attrs: { tone: "info" },
    content: [
      { type: "paragraph", content: [{ type: "text", text: "📑 自动生成的大纲", marks: [{ type: "strong" }] }] },
      ...outlineParagraphs,
    ],
  };

  return {
    ...jsonDoc,
    content: [outlineBlock, ...(jsonDoc.content ?? [])],
  };
}

/**
 * 对外统一入口。真实接入时，把 impl 换成你自己的网络请求实现即可，
 * Editor / App 层完全不需要改动。
 */
export async function applyAIPatch(jsonDoc: any, instruction = "生成大纲"): Promise<any> {
  const impl: AIPatchFn = localOutlinePatch;
  return impl(jsonDoc, instruction);
}
