import Handsontable from "handsontable";
import "handsontable/dist/handsontable.full.min.css";

const uploadButton = document.getElementById("uploadExcel");
const generateButton = document.getElementById("generateFiles");
const searchButton = document.getElementById("search-button");
const searchContatiner = document.getElementById("search-contatiner");
const searchInput = document.getElementById("search-term");

let highlightedRows = new Set();

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
      height: "700px",
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
      cells: function (row, col) {
        const cellProperties = {};
        cellProperties.renderer = customRenderer; // 사용자 정의 렌더러 사용
        return cellProperties;
      },
    });

    generateButton.disabled = false; // 파일 생성 버튼 활성화
    searchContatiner.classList.add("show");
  } else {
    alert(result.message);
  }
});

// 2. 파일 생성 및 다운로드
generateButton.addEventListener("click", async () => {
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

const customRenderer = function (
  instance,
  td,
  row,
  col,
  prop,
  value,
  cellProperties
) {
  Handsontable.renderers.TextRenderer.apply(this, arguments); // 기본 텍스트 렌더러
  if (highlightedRows.has(row)) {
    td.style.backgroundColor = "lightblue"; // 강조 스타일 적용
  } else {
    td.style.backgroundColor = ""; // 강조 스타일 적용
  }
};

searchInput.addEventListener("focus", () => {
  clearHighlights();
  searchInput.value = ""; // 입력값 초기화
});

// 엔터 키 이벤트 추가
searchInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    searchInput.blur();
    searchButton.click(); // 엔터 키를 누르면 검색 버튼 클릭
  }
});

searchButton.addEventListener("click", () => {
  const searchTerm = document.getElementById("search-term").value.trim();

  if (!searchTerm) {
    console.log("Please enter both column name and search term.");
    return;
  }

  clearHighlights();

  const data = hot.getData();
  const matchedRows = data
    .map((row, rowIndex) => ({ row, rowIndex }))
    .filter((item) =>
      item.row.some((cell) => String(cell).includes(searchTerm))
    );

  if (matchedRows.length > 0) {
    matchedRows.forEach((match) => {
      hot.scrollViewportTo(match.rowIndex, 0);
      highlightRow(match.rowIndex);
    });
  } else {
    console.log(`No match found for "${searchTerm}".`);
  }
});

function highlightRow(rowIndex) {
  highlightedRows.add(rowIndex); // 강조 표시할 행을 저장
  hot.render(); // 렌더러를 다시 실행
}

function clearHighlights() {
  highlightedRows.clear(); // 강조된 모든 행 제거
  hot.render(); // 다시 렌더링하여 스타일 초기화
}
