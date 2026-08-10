/**
 * InkFlow — Electron 预加载脚本
 * ------------------------------------------------------------
 * 通过 contextBridge 把一组白名单 IPC 包成 window.inkflow，
 * 渲染进程只能拿到这几个方法，无法直接触达 Node / Electron 全量能力。
 *
 * 桥接 API（与 src/types/electron.d.ts 保持一致）：
 *   inkflow.isElectron  : boolean —— 渲染进程用它判断是否运行在 Electron 里
 *   inkflow.openDocument()       -> Promise<{name,filePath,data,size} | null | {error}>
 *   inkflow.saveDocument({defaultName,data,saveAs}) -> Promise<{filePath,name} | null>
 *   inkflow.setCurrentFile(filePath) -> Promise<void>
 *   inkflow.onMenu(cb)            : 订阅原生菜单/快捷键动作（new|open|save|saveAs）
 *   inkflow.onOpenResult(cb)      : 订阅从命令行/文件关联打开的 .docx
 */
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('inkflow', {
  isElectron: true,

  openDocument: () => ipcRenderer.invoke('inkflow:open'),

  saveDocument: (payload) => ipcRenderer.invoke('inkflow:save', payload),

  setCurrentFile: (filePath) =>
    ipcRenderer.invoke('inkflow:set-current-file', filePath),

  /** 订阅原生菜单动作：返回一个取消订阅函数 */
  onMenu: (callback) => {
    const listener = (_e, action) => callback(action);
    ipcRenderer.on('inkflow:menu', listener);
    return () => ipcRenderer.removeListener('inkflow:menu', listener);
  },

  /** 订阅从命令行/文件关联打开的 .docx */
  onOpenResult: (callback) => {
    const listener = (_e, result) => callback(result);
    ipcRenderer.on('inkflow:open-result', listener);
    return () => ipcRenderer.removeListener('inkflow:open-result', listener);
  },
});
