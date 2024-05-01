import "handsontable/dist/handsontable.full.min.css";
import { registerAllModules } from "handsontable/registry";
import { exportXLSX } from "./functions/export-function";
import Table from "./components/Table";
import { useEffect, useState } from "react";
import { v4 as uuidv4 } from "uuid";
import SheetTab from "./components/SheetTab";
import TableV1 from "./components/TableV1";
import { exportXLSXV1 } from "./functions/export-function-v1";

registerAllModules();

function App() {
  const [tabs, setTabs] = useState(
    JSON.parse(localStorage.getItem("tabs")) ?? []
  );
  const [currentTab, setCurrentTab] = useState(
    JSON.parse(localStorage.getItem("currentTab")) ?? null
  );

  useEffect(() => {
    localStorage.setItem("tabs", JSON.stringify(tabs));
  }, [tabs]);

  useEffect(() => {
    localStorage.setItem("currentTab", JSON.stringify(currentTab));
  }, [currentTab]);

  const handleDeleteTab = (id) => {
    setTabs((tabs) => tabs.filter((tab) => tab.id !== id));
  };

  const handleOnUpdate = (updatedTab, title) => {
    setTabs((tabs) =>
      tabs.map((tab) => {
        if (updatedTab.id === tab.id) {
          tab.title = title;
        }
        return tab;
      })
    );
  };

  return (
    <div>
      {currentTab && (
        <TableV1
          tab={currentTab}
          onSave={(updatedTab, data) => {
            setCurrentTab((prev) => ({ ...prev, data }));
            setTabs((tabs) =>
              tabs.map((tab) => {
                if (tab.id === updatedTab?.id) {
                  tab.table = data;
                }
                return tab;
              })
            );
          }}
        />
      )}
      {/* {currentTab && (
        <Table
          tab={currentTab}
          onSave={(updatedTab, data) => {
            setCurrentTab((prev) => ({ ...prev, data }));
            setTabs((tabs) =>
              tabs.map((tab) => {
                if (tab.id === updatedTab?.id) {
                  tab.table = data;
                }
                return tab;
              })
            );
          }}
        />
      )} */}

      <div style={{ display: "flex", gap: "10px", padding: "10px" }}>
        <button
          onClick={() => {
            const tab = { id: uuidv4(), title: "New Sheet", table: [[]] };
            setTabs((tabs) => [...tabs, tab]);
            setCurrentTab(tab);
          }}
        >
          Add Tab
        </button>
        <button
          onClick={() => {
            // exportXLSX();
            exportXLSXV1();
          }}
        >
          Export me
        </button>
      </div>

      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          padding: "10px",
          gap: "10px",
        }}
      >
        {tabs.map((tab) => {
          return (
            <SheetTab
              key={tab.id}
              onClick={() => {
                setCurrentTab((prev) => tab);
              }}
              tab={tab}
              isActive={tab.id === currentTab?.id}
              onDelete={handleDeleteTab}
              onUpdate={handleOnUpdate}
            />
          );
        })}
      </div>
    </div>
  );
}

export default App;
