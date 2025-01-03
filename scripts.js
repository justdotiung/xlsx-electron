import Handsontable from "handsontable";
import "handsontable/dist/handsontable.full.min.css";

const uploadButton = document.getElementById("uploadExcel");
const generateButton = document.getElementById("generateFiles");

// let processedData = null;
let hot = null;
const container = document.getElementById("hot-container");

// 1. 엑셀 파일 업로드 및 데이터 가공
uploadButton.addEventListener("click", async () => {
  const result = await window.api.readExcelFile();

  if (result.status === "success") {
    const rawData = result.data;

    const processedData = rawData.map((data, _) =>
      data.reduce((a, c) => {
        a.push(c.v ? c.v : "");
        return a;
      }, [])
    );

    const headers = processedData.shift();
    hot = new Handsontable(container, {
      data: processedData,
      rowHeaders: true,
      colHeaders: headers,
      width: "100%",
      height: "400px",
      licenseKey: "non-commercial-and-evaluation",
      afterChange: function (changes, source) {
        if (changes) {
          changes.forEach(([row, prop, oldValue, newValue]) => {
            // console.log(row, prop, oldValue, newValue);
            // 변경된 셀에 클래스를 추가하거나 스타일을 변경
            const cell = hot.getCell(row, hot.propToCol(prop)); // 셀 요소 가져오기
            cell.classList.add("changed-cell"); // 변경된 셀에 'changed-cell' 클래스 추가
          });
        }
      },
    });

    // console.log(processedData);

    generateButton.disabled = false; // 파일 생성 버튼 활성화
  } else {
    alert(result.message);
  }
});

// 2. 파일 생성 및 다운로드
generateButton.addEventListener("click", async () => {
  // console.log(processedData);
  // if (!processedData) return;
  const data = hot.getData();
  data.unshift(hot.getColHeader());

  const result = await window.api.generateFiles(data);

  if (result.status === "success") {
    alert(
      `Files saved successfully:\nExcel: ${result.paths.excelPath}\nJSON: ${result.paths.jsonPath}`
    );
  } else {
    alert(result.message);
  }
});
