const { contextBridge, ipcRenderer } = require('electron');
const ExcelJS = require('exceljs');

console.log(ExcelJS);

contextBridge.exposeInMainWorld('electron', {
    ipcRenderer: {
        invoke: (channel, ...args) => ipcRenderer.invoke(channel, ...args)
    },
    ExcelJS: ExcelJS
});