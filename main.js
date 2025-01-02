const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const { readExcelFile, generateExcelFile } = require("./excelUtils"); // 유틸리티 파일 import
const path = require("path");
const fs = require("fs");

let mainWindow;
app.on("ready", () => {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, "src", "preload.js"),
      contextIsolation: true,
      // webSecurity: false,
    },
  });

  mainWindow.loadFile("index.html");
  mainWindow.webContents.openDevTools();
});

// 1. 엑셀 파일 읽기
ipcMain.handle("read-excel", async () => {
  const filePath = dialog.showOpenDialogSync(mainWindow, {
    title: "Select an Excel File",
    filters: [{ name: "Excel Files", extensions: ["xlsx", "xls"] }],
    properties: ["openFile"],
  });

  if (filePath && filePath[0]) {
    try {
      const json = readExcelFile(filePath[0]);
      // console.log(json);
      // const workbook = xlsx.readFile(filePath[0]);
      // const sheetName = workbook.SheetNames[0];
      // const worksheet = workbook.Sheets[sheetName];

      // const jsonData = xlsx.utils.sheet_to_json(worksheet); // 엑셀을 JSON으로 변환
      // return { status: "success", data: jsonData, filePath: filePath[0] };
      return { status: "success", data: json, filePath: filePath[0] };
    } catch (error) {
      return { status: "error", message: "Failed to read Excel file" };
    }
  } else {
    return { status: "error", message: "File selection canceled" };
  }
});

// 파일 이름에 번호를 붙여 저장하는 함수
function getUniqueFileName(basePath) {
  let counter = 1;
  let filePath = basePath;

  // 파일이 존재할 때마다 숫자를 추가하여 고유한 파일 경로를 생성
  while (fs.existsSync(filePath)) {
    const parsedPath = path.parse(basePath); // 경로와 확장자를 분리
    filePath = path.join(
      parsedPath.dir,
      `${parsedPath.name}(${counter})${parsedPath.ext}`
    );
    counter++; // 숫자 증가
  }

  return filePath; // 고유한 파일 경로 반환
}

// 2. 새로운 엑셀 및 JSON 파일 생성
ipcMain.handle("generate-files", async (event, processedData) => {
  const saveDir = dialog.showOpenDialogSync(mainWindow, {
    title: "Select Folder to Save Files",
    properties: ["openDirectory"],
  });

  if (saveDir && saveDir[0]) {
    try {
      const savePath = saveDir[0];

      // 새로운 엑셀 파일 생성
      // const workbook = xlsx.utils.book_new();
      // const worksheet = xlsx.utils.json_to_sheet(processedData); // JSON → 엑셀 변환
      // xlsx.utils.book_append_sheet(workbook, worksheet, "Processed Data");
      const excelPath = path.join(savePath, "processed_data.xlsx");

      // 이미 존재하는 파일이름이라면 번호를 붙여서 저장
      const uniqueExcelPath = getUniqueFileName(excelPath);
      generateExcelFile(processedData, uniqueExcelPath);
      // xlsx.writeFile(workbook, uniqueExcelPath);

      // JSON 파일 생성
      const jsonPath = path.join(savePath, "processed_data.json");
      const uniqueJsonPath = getUniqueFileName(jsonPath);
      fs.writeFileSync(uniqueJsonPath, JSON.stringify(processedData, null, 2), {
        encoding: "utf-8",
      });

      return {
        status: "success",
        message: "Files saved successfully",
        paths: { excelPath: uniqueExcelPath, jsonPath: uniqueJsonPath },
      };

      // return { status: 'success', message: 'Files saved successfully', paths: { excelPath, jsonPath } };
    } catch (error) {
      return { status: "error", message: "Failed to save files" };
    }
  } else {
    return { status: "error", message: "Save operation canceled" };
  }
});
