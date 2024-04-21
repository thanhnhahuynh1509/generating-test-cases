import dayjs from "dayjs";
import { Workbook } from "exceljs";

const PIXELS_PER_EXCEL_WIDTH_UNIT = 7.5;

export const exportXLSX = () => {
  const tabs = JSON.parse(localStorage.getItem("tabs"));
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet("My Sheet");

  let rowTabTitleIndex = 1;
  for (const tab of tabs) {
    const table = tab.table;
    const hasData = table.some((row) => row.some((el) => !!el?.trim()));
    if (hasData) {
      const [titles, data] = generateTitlesAndData(table, table[0].length);
      createTabTitle(sheet, tab, titles, rowTabTitleIndex);
      createTableTitles(sheet, titles, rowTabTitleIndex);
      createCellData(sheet, data, rowTabTitleIndex);
      for (let i = 1; i <= table[0].length; i++) {
        if (sheet.getColumn(i)) {
          sheet.getColumn(i).numFmt = "@";
        }
      }
      rowTabTitleIndex += data.length + 2;
    }
  }

  alignSheet(sheet);

  downloadFile(workbook);
};

function generateTitlesAndData(rows, totalCols) {
  const length = getRowsLength(rows);

  const titles = createComponentTitles(rows, length);
  const components = createComponents(rows, length);
  const entities = createEntities(rows, length);
  const testCases = createTestCases(rows, length, totalCols);
  const mapTest = createMapTest(components, entities, testCases);

  const data = [];

  for (const test of mapTest) {
    const eachCase = {};
    for (const title of titles) {
      let values = test.get(title);
      if (!values?.length) {
        values = [""];
      } else if (values?.length > 1) {
        values = values?.map((val) => `- ${val}`);
      }
      eachCase[title] = values.join("\n");
    }
    data.push(eachCase);
  }

  return [titles, data];
}

function createTabTitle(sheet, tab, titles, rowTabTitleIndex) {
  const tabTitle = tab.title;
  sheet.mergeCells(
    `A${rowTabTitleIndex}:${getNextCharAt(
      "A",
      titles.length - 1
    )}${rowTabTitleIndex}`
  );
  sheet.getCell(`A${rowTabTitleIndex}`).value = tabTitle;
  styleSheetTabTitle(sheet, rowTabTitleIndex);
}

function createTableTitles(sheet, titles, rowTabTitleIndex) {
  for (let i = 0; i < titles.length; i++) {
    sheet.getCell(`${getNextCharAt("A", i)}${rowTabTitleIndex + 1}`).value =
      titles[i];
  }
  styleSheetTitle(sheet, rowTabTitleIndex + 1);
}

function createCellData(sheet, data, rowTabTitleIndex) {
  data.forEach((row, i) => {
    Object.values(row).forEach((val, colIdx) => {
      const cell = sheet.getCell(
        `${getNextCharAt("A", colIdx)}${rowTabTitleIndex + 2 + i}`
      );
      cell.numFmt = "@";
      cell.value = `${val}`;
    });
  });
}

function getNextCharAt(char, index) {
  return String.fromCharCode(char.charCodeAt(0) + index);
}

function autoSize(sheet, fromRow) {
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return;
  }

  const maxColumnLengths = [];
  sheet.eachRow((row, rowNum) => {
    if (rowNum < fromRow) {
      return;
    }

    row.eachCell((cell, num) => {
      if (typeof cell.value === "string") {
        if (maxColumnLengths[num] === undefined) {
          maxColumnLengths[num] = 0;
        }

        const fontSize = cell.font && cell.font.size ? cell.font.size : 11;
        ctx.font = `${fontSize}pt Arial`;
        const metrics = ctx.measureText(cell.value);
        const cellWidth = metrics.width;

        maxColumnLengths[num] = Math.max(maxColumnLengths[num], cellWidth);
      }
    });
  });

  for (let i = 1; i <= sheet.columnCount; i++) {
    const col = sheet.getColumn(i);
    const width = maxColumnLengths[i];
    if (width) {
      col.width = width / PIXELS_PER_EXCEL_WIDTH_UNIT + 1;
    }
  }
}

function getRowsLength(rows) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const entity = row[1];
    if (entity === undefined || entity === null) {
      return i;
    }
  }
}

function createComponentTitles(rows, length) {
  const titles = [];
  for (let i = 0; i < length; i++) {
    if (rows[i][0]) {
      titles.push(rows[i][0]);
    }
  }
  return titles;
}

function createComponents(rows, length) {
  const components = [];
  let prevComponent = null;
  for (let i = 0; i < length; i++) {
    if (rows[i][0]) {
      prevComponent = rows[i][0];
    }
    components.push(prevComponent);
  }
  return components;
}

function createEntities(rows, length) {
  const entities = [];
  for (let i = 0; i < length; i++) {
    entities.push(rows[i][1]);
  }
  return entities;
}

function createTestCases(rows, length, totalCols) {
  const testCases = [];
  for (let colIdx = 2; colIdx < totalCols; colIdx++) {
    const testCase = [];
    for (let rowIdx = 0; rowIdx < length; rowIdx++) {
      if (rows[rowIdx][colIdx]?.trim()) {
        testCase.push(rowIdx);
      }
    }
    if (testCase.length) {
      testCases.push(testCase);
    }
  }
  return testCases;
}

function createMapTest(components, entities, testCases) {
  const result = [];
  for (const testCase of testCases) {
    const mapTest = new Map();
    for (const i of testCase) {
      const component = components[i];
      const entity = entities[i];
      if (!mapTest.get(component)) {
        mapTest.set(component, []);
      }
      const elements = mapTest.get(component);
      elements.push(entity);
      mapTest.set(component, elements);
    }
    result.push(mapTest);
  }
  return result;
}

function styleSheetTabTitle(sheet, idx) {
  sheet.getRow(idx).fill = {
    type: "pattern",
    pattern: "darkGray",
    fgColor: { argb: "ff0f435f" },
  };

  sheet.getRow(idx).font = {
    bold: true,
    color: { argb: "ffffff" },
    size: "11",
  };
}

function styleSheetTitle(sheet, idx) {
  const row = sheet.getRow(idx);
  row.fill = {
    type: "pattern",
    pattern: "darkGray",
    fgColor: { argb: "417c91" },
  };

  row.font = {
    bold: true,
    color: { argb: "ffffff" },
    size: "11",
  };
}

function alignSheet(sheet) {
  autoSize(sheet, 0);
  for (let i = 2; i <= sheet.actualRowCount; i++) {
    const row = sheet.getRow(i);
    row.alignment = { vertical: "top" };
  }
}

function downloadFile(workbook) {
  const fileName = `${dayjs(Date.now()).format("DD/MM/YYYY_HH[h]mm[p]")}.xlsx`;
  workbook.xlsx.writeBuffer().then((data) => {
    const blob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheet.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    window.URL.revokeObjectURL(url);
  });
}
