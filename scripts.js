const uploadButton = document.getElementById("uploadExcel");
const outputArea = document.getElementById("output");
const generateButton = document.getElementById("generateFiles");

let processedData = null;

// 1. 엑셀 파일 업로드 및 데이터 가공
uploadButton.addEventListener("click", async () => {
  const result = await window.api.readExcelFile();

  if (result.status === "success") {
    const rawData = result.data;

    // 예시: 데이터를 간단히 가공 (열 추가)
    processedData = rawData;
    // processedData = rawData.map((row, index) => ({
    //   ...row,
    //   RowNumber: index + 1, // 예시로 행 번호 추가
    // }));

    outputArea.value = JSON.stringify(rawData, null, 2); // JSON 출력
    generateButton.disabled = false; // 파일 생성 버튼 활성화
  } else {
    alert(result.message);
  }
});

// 2. 파일 생성 및 다운로드
generateButton.addEventListener("click", async () => {
  console.log(processedData);
  if (!processedData) return;

  const result = await window.api.generateFiles(processedData);

  if (result.status === "success") {
    alert(
      `Files saved successfully:\nExcel: ${result.paths.excelPath}\nJSON: ${result.paths.jsonPath}`
    );
  } else {
    alert(result.message);
  }
});
