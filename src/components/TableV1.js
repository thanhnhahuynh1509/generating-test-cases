import { useEffect, useRef, useState } from "react";
import { HotTable } from "@handsontable/react";

function TableV1({ tab, onSave }) {
  const hotRef = useRef();
  const [height, setHeight] = useState(window.innerHeight);
  const [renderedData, setRenderData] = useState(tab?.table);

  useEffect(() => {
    const resizeEvent = window.addEventListener("resize", () => {
      setHeight(window.innerHeight);
    });

    return () => {
      window.removeEventListener("resize", resizeEvent);
    };
  }, []);

  useEffect(() => {
    setRenderData((prev) => tab?.table ?? [[]]);
  }, [tab]);

  return (
    <HotTable
      data={renderedData}
      rowHeaders={true}
      colHeaders={["Data"]}
      height={`${height - 200}`}
      autoWrapRow={true}
      autoWrapCol={true}
      minCols={100}
      minRows={100}
      fixedColumnsStart={1}
      fixedRowsTop={2}
      manualColumnResize={true}
      manualColumnMove
      manualRowMove
      collapsibleColumns={true}
      contextMenu={true}
      ref={hotRef}
      afterChange={(changes, _source) => {
        if (!changes?.length) {
          return;
        } else {
          const prevContent = changes[0][2];
          const updatedContent = changes[0][3];
          if (prevContent === updatedContent) {
            return;
          }
        }
        if (onSave) {
          console.log(hotRef.current.hotInstance.getData());
          onSave(tab, hotRef.current.hotInstance.getData());
        }
      }}
      cells={function (row, col, _prop) {
        const data = this.instance.getDataAtCell(row, col);
        if (col === 0) {
          if (
            data &&
            ["[PRIORITY]", "[RESOURCE]", "[RESULT]"].includes(
              data.toUpperCase()
            )
          ) {
            for (let i = 0; i < this.instance.countCols(); i++) {
              this.instance.setCellMeta(row, i, "className", ` bg-title`);
            }
          } else if (isEnclosedInBrackets(data?.trim())) {
            this.instance.setCellMeta(row, 0, "className", ` text-bold`);
          }
        }
      }}
      licenseKey="non-commercial-and-evaluation"
    />
  );
}

function isEnclosedInBrackets(str) {
  const regex = /^<[^<>]+>$/;
  return regex.test(str);
}

export default TableV1;
