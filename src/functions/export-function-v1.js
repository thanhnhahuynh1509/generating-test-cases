import dayjs from "dayjs";
import { Workbook } from "exceljs";

const PIXELS_PER_EXCEL_WIDTH_UNIT = 7.5;

export const exportXLSXV1 = () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet("My Sheet");

  const tabs = JSON.parse(localStorage.getItem("tabs"));
  console.log(tabs);
  console.log(tabs[0].table);
  let rowIdx = 0;
  for (const tab of tabs) {
    const table = tab.table;
    createTabTitle(sheet, tab, rowIdx);
    rowIdx += 1;

    for (let i = 1; i <= table[0].length; i++) {
      if (sheet.getColumn(i)) {
        sheet.getColumn(i).numFmt = "@";
        sheet.getColumn(i).alignment = {
          vertical: "top",
          horizontal: "left",
          wrapText: true,
        };
      }
    }

    const hasData = table.some((row) => row.some((el) => !!el?.trim()));

    if (hasData) {
      const priorityIndex = table.findIndex((value) =>
        value.find((value) => value?.trim()?.toLowerCase() === "[priority]")
      );
      const resourceIndex = table.findIndex((value) =>
        value.find((value) => value?.trim()?.toLowerCase() === "[resource]")
      );
      const resultIndex = table.findIndex((value) =>
        value.find((value) => value?.trim()?.toLowerCase() === "[result]")
      );

      const priorityData = [
        "",
        ...table[priorityIndex + 1].filter((data) => !!data),
      ];

      const testCases = [];
      for (let col = 1; col < priorityData.length; col++) {
        const resourcesData = [];
        for (let row = resourceIndex + 1; row < resultIndex; row++) {
          const title = table[row][0];
          const hasData = !!table[row][col]?.trim();
          if (hasData) {
            resourcesData.push(title);
          }
        }

        const resultsData = [];
        for (let row = resultIndex + 1; row < table[0].length; row++) {
          const title = table[row][0];
          const hasData = !!table[row][col]?.trim();
          if (hasData) {
            resultsData.push(title);
          }
        }

        testCases.push({
          resourcesData,
          resultsData,
          priority: getPriorityLabel(priorityData[col]),
        });
      }

      for (const testCase of testCases) {
        const { resourcesData, resultsData, priority } = testCase;

        createCell(sheet, rowIdx, 0, priority);
        createCell(
          sheet,
          rowIdx,
          1,
          resourcesData
            .map((resource, idx) => `${idx + 1}. ${resource}`)
            .reduce((acc, data) => acc + data + "\n", "")
        );
        createCell(
          sheet,
          rowIdx,
          2,
          resultsData
            .map((resource) => `- ${resource}`)
            .reduce((acc, data) => acc + data + "\n", "")
        );

        rowIdx += 1;
      }
    }
  }
  alignSheet(sheet);
  downloadFile(workbook);
};

function getPriorityLabel(priority) {
  switch (priority.toLowerCase()) {
    case "l":
      return "Low";
    case "m":
      return "Medium";
    case "h":
      return "High";
    default:
      return "Medium";
  }
}

function convertToStrings(dataSource, isOrder) {
  return dataSource.map((item, idx) => {
    const strings = [];
    for (const element of Object.keys(item)) {
      const data = item[element];
      let value = ``;
      if (isOrder) {
        value += `${idx + 1}. `;
      }
      if (element) {
        value += `${element}: `;
      }

      if (data?.length > 1) {
        if (element) {
          value += "\n" + data.map((val) => `+ ${val}`).join("\n");
        } else {
          value += data.map((val) => `${val}`).join("\n");
        }
      } else {
        value += data.join("");
      }

      strings.push(value);
    }

    return strings;
  });
}

function convertToRichtext(dataSource, isOrder) {
  let value = {
    richText: [],
  };
  dataSource.forEach((item, idx) => {
    for (const element of Object.keys(item)) {
      const data = item[element];

      if (isOrder) {
        value.richText.push({ text: `${idx + 1}. ` });
      }
      if (element) {
        value.richText.push({ text: `${element}: `, font: { bold: true } });
      }

      if (data?.length > 1) {
        if (element) {
          value.richText.push({
            text: "\n" + data.map((val) => `+ ${val}`).join("\n"),
          });
        } else {
          value.richText.push({
            text: data.map((val) => `${val}`).join("\n"),
          });
        }
      } else {
        value.richText.push({
          text: data.join(""),
        });
      }

      if (idx < dataSource.length - 1) {
        value.richText.push({ text: "\n" });
      }
    }
  });

  return value;
}

function createCell(sheet, rowIdx, colIdx, value) {
  const cell = sheet.getCell(`${getNextCharAt("A", colIdx)}${rowIdx}`);
  cell.numFmt = "@";
  cell.value = value;

  return cell;
}

function getNextCharAt(char, index) {
  return String.fromCharCode(char.charCodeAt(0) + index);
}

function alignSheet(sheet) {
  autoSize(sheet, 0);
  for (let i = 2; i <= sheet.actualRowCount; i++) {
    const row = sheet.getRow(i);
    row.alignment = { vertical: "top" };
  }
}

function createTabTitle(sheet, tab, rowTabTitleIndex) {
  const tabTitle = tab.title;
  sheet.mergeCells(
    `A${rowTabTitleIndex}:${getNextCharAt("A", 3)}${rowTabTitleIndex}`
  );
  sheet.getCell(`A${rowTabTitleIndex}`).value = tabTitle;
  styleSheetTabTitle(sheet, rowTabTitleIndex);
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
