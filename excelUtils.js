const xlsx = require("xlsx");

// 열 번호를 열 이름으로 변환하는 유틸리티 함수
function colToLetter(col) {
  let letter = "";
  while (col >= 0) {
    letter = String.fromCharCode((col % 26) + 65) + letter;
    col = Math.floor(col / 26) - 1;
  }
  return letter;
}

// 특정 셀 범위를 묶어서 읽는 함수
// 날짜 처리
function getCellRangeData(sheet, startRow, endRow, startCol, endCol) {
  const data = [];

  for (let row = startRow; row <= endRow; row += 2) {
    // 2행씩 묶음 처리
    const group = [];

    // 첫 번째 줄 (row)
    const row1 = [];
    for (let col = startCol; col <= endCol; col++) {
      if (col === 88) {
        // 날짜 비고처리란 제거
        continue;
      }
      const cellAddress = String.fromCharCode(col) + row; // ex: I8
      const cellData = sheet[cellAddress];
      row1.push(
        cellData
          ? { v: cellData.v, key: cellAddress }
          : { v: 0, key: cellAddress }
      );
    }

    // 두 번째 줄 (row + 1)
    const row2 = [];
    for (let col = startCol; col <= endCol; col++) {
      const cellAddress = String.fromCharCode(col) + (row + 1); // ex: I9
      const cellData = sheet[cellAddress];
      row2.push(
        cellData
          ? { v: cellData.v, key: cellAddress }
          : { v: 0, key: cellAddress }
      );
    }

    // 두 줄을 하나의 묶음으로 추가
    group.push(...row1.concat(row2));
    data.push(group);
  }

  return data;
}

// 데이터 추출 함수
function getRowData2(
  sheet,
  startRow,
  endRow,
  startCol,
  endCol,
  separator,
  separator1
) {
  const prvData = [];

  const nameData = [];
  for (let col = startCol; col <= endCol; col++) {
    const colLetter = colToLetter(col); // 숫자에서 알파벳으로 변환
    const cellAddress = colLetter + 6; // 셀 주소
    const cellData = sheet[cellAddress];
    //   rowData.push(cellData ? cellData.v : null);
    nameData.push(
      cellData
        ? typeof cellData.v === "string"
          ? cellData.v.replaceAll(" ", "")
          : cellData.v
        : null
    );
  }

  const workingdata = getCellRangeData(
    sheet,
    startRow,
    endRow,
    separator,
    separator1
  );
  // 각 행에 대해 데이터 추출

  for (let row = startRow; row <= endRow; row += 2) {
    const rowData = [];
    let i = 0;
    let isAppend = false;
    for (let col = startCol; col <= endCol; col++) {
      const colLetter = colToLetter(col);
      if (8 <= col && 23 >= col) {
        if (!isAppend) {
          isAppend = true;
          const a = workingdata.shift();

          if (a) rowData.push(...a);
        }

        continue;
      }

      const cellAddress = colLetter + row; // 셀 주소
      const cellData = sheet[cellAddress];

      rowData.push(
        cellData
          ? {
              v:
                nameData[i] === "주민등록번호"
                  ? formatString(cellData.v.toString().replace("-", ""))
                  : cellData.v,
              key: nameData[i] === "오류" ? "국적" : nameData[i],
            }
          : {
              v: "null확인필요",
              key: nameData[i] === "오류" ? "국적" : nameData[i],
            }
      );

      i++;
    }
    prvData.push(rowData);
  }

  function formatString(input) {
    // 앞 6자리와 뒤 7자리를 분리하여 포맷팅
    return `${input.slice(0, 6)}-${input.slice(6)}`;
  }

  const first = nameData.splice(0, 8);
  const last = nameData.splice(24 - 8);

  const mid = Array.from({ length: 31 }, (v, i) => i + 1);

  const full = first.concat(mid, last);

  const tags = full.map((v) => {
    let newValue = v === "오류" ? "국적" : v;

    if (typeof v === "number") newValue = `${v}일`;

    return { v: newValue };
  });
  prvData.unshift(tags);
  return prvData;
}

function readExcelFile(filePath) {
  const workbook = xlsx.readFile(filePath);

  // 첫 번째 시트 이름 가져오기
  const sheetName = workbook.SheetNames[0];

  // 첫 번째 시트 데이터 가져오기
  const sheet = workbook.Sheets[sheetName];
  const sheetRange = sheet["!ref"];
  const range = xlsx.utils.decode_range(sheetRange);
  const lastRow = range.e.r + 1; // 0부터 시작하므로 +1

  // 'I'부터 'X'까지, 8행부터 시작
  const startRow = 8; // 시작 행
  const startCol = "I".charCodeAt(0); // 시작 열 (I)
  const endCol = "X".charCodeAt(0); // 끝 열 (X)

  // 열 범위 설정
  const startCol2 = 0; // 'A' -> 0부터 시작
  const endCol2 = 35; // 'AJ' -> 35까지 (0부터 시작하므로 AJ는 35번째 열)

  const result = getCellRangeData(sheet, startRow, lastRow, startCol, endCol);
  const result2 = getRowData2(
    sheet,
    startRow,
    lastRow,
    startCol2,
    endCol2,
    startCol,
    endCol
  );

  return result2;
}

function generateExcelFile(processedData, filePath) {
  const workbook = xlsx.utils.book_new();

  const worksheet = xlsx.utils.json_to_sheet(processedData); // JSON → 엑셀 변환
  xlsx.utils.book_append_sheet(workbook, worksheet, "Processed Data");
  xlsx.writeFile(workbook, filePath);
}

module.exports = { readExcelFile, generateExcelFile };
