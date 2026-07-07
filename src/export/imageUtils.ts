const MAX_WIDTH_PX = 560;

export function parseDataUrl(
  src: string
): { base64: string; mime: string } | null {
  const match = /^data:([^;]+);base64,(.*)$/s.exec(src);
  if (!match) return null;
  return { mime: match[1], base64: match[2] };
}

export function docxImageType(mime: string): 'jpg' | 'png' | 'gif' | 'bmp' {
  if (mime.includes('jpeg') || mime.includes('jpg')) return 'jpg';
  if (mime.includes('gif')) return 'gif';
  if (mime.includes('bmp')) return 'bmp';
  return 'png'; // 兜底：docx.js 只认 jpg/png/gif/bmp
}

/** 探测图片真实尺寸，并按最大宽度等比缩放，避免导出的 DOCX 图片撑破页面 */
export function getImageDimensions(
  src: string
): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    const fallback = { width: 320, height: 200 };
    const timer = setTimeout(() => resolve(fallback), 4000);
    img.onload = () => {
      clearTimeout(timer);
      const w = img.naturalWidth || fallback.width;
      const h = img.naturalHeight || fallback.height;
      if (w <= MAX_WIDTH_PX) {
        resolve({ width: w, height: h });
      } else {
        const ratio = MAX_WIDTH_PX / w;
        resolve({ width: MAX_WIDTH_PX, height: Math.round(h * ratio) });
      }
    };
    img.onerror = () => {
      clearTimeout(timer);
      resolve(fallback);
    };
    img.src = src;
  });
}

/** 远程图片转 data URL（供无法直接以 base64 传给 docx.js 的场景使用）*/
export async function remoteImageToDataUrl(
  url: string
): Promise<string | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const blob = await res.blob();
    return await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}
