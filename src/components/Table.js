import { useEffect, useRef, useState } from "react";
import { HotTable } from "@handsontable/react";
import {
  handleBeforeChange,
  handleOnCellMouseDown,
} from "../functions/handle-event-functions";

function Table({ tab, onSave }) {
  const hotRef = useRef();
  const [height, setHeight] = useState(window.innerHeight);
  const [renderedData, setRenderData] = useState(tab?.table);
  //  ?? MOCK_DATA

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

  console.log("update")

  return (
    <HotTable
      data={renderedData}
      rowHeaders={true}
      colHeaders={["Components", "Elements"]}
      height={`${height - 200}`}
      autoWrapRow={true}
      autoWrapCol={true}
      minCols={100}
      minRows={100}
      fixedColumnsStart={2}
      manualColumnResize={true}
      manualColumnMove
      manualRowMove
      contextMenu={true}
      ref={hotRef}
      afterOnCellMouseDown={handleOnCellMouseDown(hotRef)}
      afterChange={(changes, _source) => {
        if (!changes?.length) {
          return;
        }
        if (onSave) {
          onSave(tab, hotRef.current.hotInstance.getData());
        }
      }}
      beforeChange={handleBeforeChange(hotRef)}
      cells={function (row, col, _prop) {
        const data = this.instance.getDataAtCell(row, col);
        if (data && col === 0) {
          for (let i = 0; i < this.instance.countCols(); i++) {
            this.instance.setCellMeta(row, i, "className", ` border_top`);
          }
        }
      }}
      licenseKey="non-commercial-and-evaluation" // for non-commercial use only
    />
  );
}

export default Table;
