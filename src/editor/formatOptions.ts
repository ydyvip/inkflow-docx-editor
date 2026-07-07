/** 字体列表：常见西文 + 中文字体，贴近 Word 字体下拉的常用项 */
export const FONT_FAMILIES: { label: string; value: string }[] = [
  { label: '正文默认', value: '' },
  { label: '微软雅黑', value: 'Microsoft YaHei' },
  { label: '宋体', value: 'SimSun' },
  { label: '黑体', value: 'SimHei' },
  { label: '楷体', value: 'KaiTi' },
  { label: '仿宋', value: 'FangSong' },
  { label: 'Arial', value: 'Arial' },
  { label: 'Times New Roman', value: 'Times New Roman' },
  { label: 'Georgia', value: 'Georgia' },
  { label: 'Courier New', value: 'Courier New' },
  { label: 'Calibri', value: 'Calibri' },
];

/** 字号列表（磅值），对应存储单位是半磅（pt * 2）*/
export const FONT_SIZES_PT = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48,
];

export const LINE_SPACING_OPTIONS = [
  { label: '默认行距', value: '' },
  { label: '1.0', value: '1' },
  { label: '1.15', value: '1.15' },
  { label: '1.5', value: '1.5' },
  { label: '2.0', value: '2' },
];

/** Word 高亮底色调色板（对应 OOXML w:highlight 的取值关键字）*/
export const HIGHLIGHT_OPTIONS: {
  label: string;
  value: string;
  css: string;
}[] = [
  { label: '无', value: '', css: 'transparent' },
  { label: '黄', value: 'yellow', css: '#FFFF00' },
  { label: '绿', value: 'green', css: '#00FF00' },
  { label: '青', value: 'cyan', css: '#00FFFF' },
  { label: '粉', value: 'magenta', css: '#FF00FF' },
  { label: '蓝', value: 'blue', css: '#0000FF' },
  { label: '红', value: 'red', css: '#FF0000' },
  { label: '浅灰', value: 'lightGray', css: '#D3D3D3' },
];

export const HEADING_LEVELS = [1, 2, 3, 4, 5, 6, 7, 8, 9] as const;
