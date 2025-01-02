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
        // console.log(String.fromCharCode(col) + row);
        continue;
      }
      // console.log(row);
      // if (row === 8) continue;
      const cellAddress = String.fromCharCode(col) + row; // ex: I8
      const cellData = sheet[cellAddress];
      // row1.push(cellData ? `{v: ${cellData.v}, cellData: ${cellAddress}}`: null);
      row1.push(cellData ? { v: cellData.v, key: cellAddress } : null);
    }

    // 두 번째 줄 (row + 1)
    const row2 = [];
    for (let col = startCol; col <= endCol; col++) {
      const cellAddress = String.fromCharCode(col) + (row + 1); // ex: I9
      const cellData = sheet[cellAddress];
      // row2.push(cellData ? `{v: ${cellData.v}, cellData: ${cellAddress}}`: null);
      row2.push(cellData ? { v: cellData.v, key: cellAddress } : null);
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
    //   console.log(colLetter, col);
    const cellAddress = colLetter + 6; // 셀 주소
    // console.log(cellAddress);
    const cellData = sheet[cellAddress];
    //   rowData.push(cellData ? cellData.v : null);
    // if (cellData) console.log(cellData.v);
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
  //   console.log(nameData);
  // 각 행에 대해 데이터 추출

  for (let row = startRow; row <= endRow; row += 2) {
    const rowData = [];
    let i = 0;
    let isAppend = false;
    for (let col = startCol; col <= endCol; col++) {
      const colLetter = colToLetter(col);
      if (8 <= col && 24 >= col) {
        if (!isAppend) {
          // console.log(workingdata);
          isAppend = true;
          const a = workingdata.shift();

          if (a) rowData.push(...a);
        }
        // console.log(colLetter);

        continue;
      }
      // console.log(colLetter);
      // console.log(separator, separator1, col);
      // if (25 <= col && isAppend) {
      //   console.log(colLetter, i);
      //   rowData.push(...workingdata.shift());
      //   isAppend = false;
      // }

      // const colLetter = colToLetter(col); // 숫자에서 알파벳으로 변환
      // console.log(colLetter);
      const cellAddress = colLetter + row; // 셀 주소
      // console.log(cellAddress);
      const cellData = sheet[cellAddress];
      rowData.push(
        cellData
          ? {
              v: cellData.v,
              key: nameData[i] === "오류" ? "국적" : nameData[i],
            }
          : null
      );

      i++;
    }
    // isAppend = false;
    prvData.push(rowData);
  }

  // // 각 행에 대해 데이터 추출
  // for (let row = startRow; row <= endRow; row += 2) {
  //   const rowData = [];
  //   let i = 0;
  //   for (let col = startCol; col <= endCol; col++) {
  //     const colLetter = colToLetter(col); // 숫자에서 알파벳으로 변환
  //     // console.log(colLetter, col);
  //     const cellAddress = colLetter + row; // 셀 주소
  //     // console.log(cellAddress);
  //     const cellData = sheet[cellAddress];
  //     // if (cellData && typeof nameData[i] === "number")
  //     //   console.log(cellData.v, nameData[i], cellAddress);
  //     //   rowData.push(cellData ? cellData.v : null);
  //     //   rowData.push(cellData ? `{v: ${cellData.v}, key: ${nameData[i] === '오류' ? '국적' : nameData[i]}}`: null);
  //     rowData.push(
  //       cellData
  //         ? {
  //             v: cellData.v,
  //             key: nameData[i] === "오류" ? "국적" : nameData[i],
  //           }
  //         : null
  //     );

  //     i++;
  //   }

  //   // console.log(rowData);
  //   data.push(rowData);
  // }

  const tags = [...nameData].map((v, i) => {
    let newValue = v === "오류" ? "국적" : v;
    if (i >= 8 && typeof v === "number") newValue = i - 7;

    return { v: `${newValue}일` };
  });
  // console.log(data);
  prvData.unshift(tags);
  return prvData;
}

// insertData(sheet1, 2, 0, result2);

// xlsx.writeFile(workbook1, filePath);

function readExcelFile(filePath) {
  console.log(filePath);
  const workbook = xlsx.readFile(filePath);

  // 첫 번째 시트 이름 가져오기
  const sheetName = workbook.SheetNames[0];
  // console.log(2222);

  // 첫 번째 시트 데이터 가져오기
  const sheet = workbook.Sheets[sheetName];
  const sheetRange = sheet["!ref"];
  const range = xlsx.utils.decode_range(sheetRange);
  const lastRow = range.e.r + 1; // 0부터 시작하므로 +1

  // 'I'부터 'X'까지, 8행부터 시작
  const startRow = 8; // 시작 행
  const startCol = "I".charCodeAt(0); // 시작 열 (I)
  const endCol = "X".charCodeAt(0); // 끝 열 (X)

  // console.log(111);
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
  // console.log(result);

  // result2.map((r, i) => r.push({ v: result[i], key: "일수" }));

  // insertData(sheet1, 2, 0, result2);
  // result2.unshift([{ v: "일수", key: null }]);
  return result2;
}

function generateExcelFile(processedData, filePath) {
  // 1. 기존 XLSX 파일 읽기
  // 2. 새로운 XLSX 파일 생성
  // 3. 기존 XLSX의 모든 시트를 새로운 파일에 복사
  // 4. 새로운 XLSX 파일 저장
  const workbook = xlsx.utils.book_new();
  // // workbook1.SheetNames.forEach((sheetName) => {
  // //   const sheet = workbook.Sheets[sheetName];
  // //   xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
  // // });
  // // sheet1
  // xlsx.utils.book_append_sheet(workbook, sheet1, sheetName1);
  // // const workbook1 = xlsx.readFile("Processed Data");
  // // 2. 첫 번째 시트 가져오기
  // const sheetName = workbook.SheetNames[0];
  // console.log(sheetName);
  // const sheet = workbook.Sheets[sheetName];
  // if (!sheet) {
  //   // 만약 해당 시트가 없다면 새 시트를 생성
  //   sheet = xlsx.utils.aoa_to_sheet([]);
  //   xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
  // }
  // // 기존 데이터를 읽고 병합
  // // const existingData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  // // const mergedData = [...existingData, ...processedData];
  // // // 업데이트된 데이터를 시트에 반영
  // // const updatedSheet = xlsx.utils.aoa_to_sheet(mergedData);
  // // workbook.Sheets[sheet] = updatedSheet;
  // // xlsx.writeFile(workbook1, filePath);
  // insertData(sheet1, 2, 0, processedData);
  const worksheet = xlsx.utils.json_to_sheet(processedData); // JSON → 엑셀 변환
  xlsx.utils.book_append_sheet(workbook, worksheet, "Processed Data");
  xlsx.writeFile(workbook, filePath);
  // insertData(sheet1, 2, 0, processedData);
  // xlsx.writeFile(workbook1, filePath);
}

module.exports = { readExcelFile, generateExcelFile };
