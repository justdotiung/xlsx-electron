// const path = require("path");
const { contextBridge, ipcRenderer } = require("electron");

console.log("Preload script has been loaded successfully!");

contextBridge.exposeInMainWorld("api", {
  readExcelFile: () => ipcRenderer.invoke("read-excel"),
  generateFiles: (processedData) =>
    ipcRenderer.invoke("generate-files", processedData),
});
