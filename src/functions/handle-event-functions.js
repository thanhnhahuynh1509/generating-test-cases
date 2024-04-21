export const handleOnCellMouseDown = (hotRef) => {
  return (_e, coords, _) => {
    const { row, col } = coords;
    if (row < 0 || col < 0) {
      return;
    }
    if (col > 1) {
      const dataAtCell = hotRef.current.hotInstance.getDataAtCell(row, col);
      let data = "";
      if (!dataAtCell) {
        data = "x";
      }
      hotRef.current.hotInstance.setDataAtCell(row, col, data);
    }
  };
};

const updateBorderRow = (hotRef) => {
  return (row, className) => {
    const colsLength = hotRef.current.hotInstance.countCols();
    for (let i = 0; i < colsLength; i++) {
      const cellMeta = hotRef.current.hotInstance.getCellMeta(row, i);
      hotRef.current.hotInstance.setCellMeta(
        row,
        i,
        "className",
        cellMeta.className?.replace("border_top", "") + ` ${className}`
      );
    }
  };
};

export const handleBeforeChange = (hotRef) => {
  return (changes, _source) => {
    if (!changes?.length) {
      return;
    }
    const [row, col, _, value] = changes[0];
    let className = "border_top";
    if (!value?.trim()) {
      className = className.replace("border_top", "");
    }
    if (col === 0) {
      updateBorderRow(hotRef)(row, className);
    }
  };
};

export const handleAfterChange = (hotRef) => {
  return (changes, _source) => {
    if (!changes?.length) {
      return;
    }
    localStorage.setItem(
      "tableData",
      JSON.stringify(hotRef.current.hotInstance.getData())
    );
  };
};
