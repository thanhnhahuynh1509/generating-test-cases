import { useEffect, useState } from "react";

function SheetTab({ tab, isActive, onClick, onDelete, onUpdate }) {
  const [onConfirmDelete, setOnConfirmDelete] = useState(false);
  const [title] = useState(tab?.title);
  useEffect(() => {
    const id = window.addEventListener("click", () => {
      setOnConfirmDelete(false);
    });

    return () => window.removeEventListener("click", id);
  }, []);

  return (
    <div
      key={tab.id}
      className={`card ${isActive && "active"}`}
      onClick={onClick}
      style={{ userSelect: "none", fontSize: `14px` }}
      contentEditable={isActive}
      suppressContentEditableWarning
      onInput={(e) => {
        if (onUpdate) {
          onUpdate(tab, e.currentTarget.textContent);
        }
      }}
    >
      {title}
      {!isActive && (
        <div className="toolbar">
          <div
            className="delete tool-item"
            style={{
              top: "-10px",
              right: "-5px",
              cursor: "pointer",
            }}
            onClick={(e) => {
              e.stopPropagation();
              setOnConfirmDelete((prev) => !prev);
            }}
          >
            <i className="fa-solid fa-trash"></i>
            {onConfirmDelete && (
              <>
                <div
                  className="delete tool-item-child"
                  style={{
                    top: "-15px",
                    right: "-25px",
                    cursor: "pointer",
                    zIndex: "10000",
                  }}
                  onClick={(e) => {
                    e.stopPropagation();
                    setOnConfirmDelete(false);
                  }}
                >
                  <i className="fa-solid fa-xmark" style={{ color: "red" }}></i>
                </div>
                <div
                  className="delete tool-item-child"
                  style={{
                    top: "-15px",
                    right: "25px",
                    cursor: "pointer",
                    zIndex: "10000",
                  }}
                  onClick={(e) => {
                    e.stopPropagation();
                    onDelete(tab.id);
                  }}
                >
                  <i
                    className="fa-solid fa-check"
                    style={{ color: "green" }}
                  ></i>
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

export default SheetTab;
