import React from "react";
import * as XLSX from "xlsx";

const data = [
  {
    "Column 1": "value 1",
    "Column 2": "value 1",
    "Column 3": "value 1",
    "Column 4": "value 1",
    "Column 5": "value 1",
    "Column 6": "value 1",
    "Column 7": "value 1",
    "Column 8": "value 1",
    "Column 9": "value 1",
    "Column 10": "value 1",
    "Column 11": "value 1",
    "Column 12": "value 1",
    "Column 13": "value 1",
    "Column 14": "value 1",
    "Column 15": "value 1",
    "Column 16": "value 1",
    "Column 17": "value 1",
    "Column 18": "value 1",
    "Column 19": "value 1",
    "Column 20": "value 1",
    "Column 21": "value 1",
    "Column 22": "value 1",
  },
  {
    "Column 1": "value 2",
    "Column 2": "value 2",
    "Column 3": "value 2",
    "Column 4": "value 2",
    "Column 5": "value 2",
    "Column 6": "value 2",
    "Column 7": "value 2",
    "Column 8": "value 2",
    "Column 9": "value 2",
    "Column 10": "value 2",
    "Column 11": "value 2",
    "Column 12": "value 2",
    "Column 13": "value 2",
    "Column 14": "value 2",
    "Column 15": "value 2",
    "Column 16": "value 2",
    "Column 17": "value 2",
    "Column 18": "value 2",
    "Column 19": "value 2",
    "Column 20": "value 2",
    "Column 21": "value 2",
    "Column 22": "value 2",
  },
  {
    "Column 1": "value 3",
    "Column 2": "value 3",
    "Column 3": "value 3",
    "Column 4": "value 3",
    "Column 5": "value 3",
    "Column 6": "value 3",
    "Column 7": "value 3",
    "Column 8": "value 3",
    "Column 9": "value 3",
    "Column 10": "value 3",
    "Column 11": "value 3",
    "Column 12": "value 3",
    "Column 13": "value 3",
    "Column 14": "value 3",
    "Column 15": "value 3",
    "Column 16": "value 3",
    "Column 17": "value 3",
    "Column 18": "value 3",
    "Column 19": "value 3",
    "Column 20": "value 3",
    "Column 21": "value 3",
    "Column 22": "value 3",
  },
  {
    "Column 1": "value 4",
    "Column 2": "value 4",
    "Column 3": "value 4",
    "Column 4": "value 4",
    "Column 5": "value 4",
    "Column 6": "value 4",
    "Column 7": "value 4",
    "Column 8": "value 4",
    "Column 9": "value 4",
    "Column 10": "value 4",
    "Column 11": "value 4",
    "Column 12": "value 4",
    "Column 13": "value 4",
    "Column 14": "value 4",
    "Column 15": "value 4",
    "Column 16": "value 4",
    "Column 17": "value 4",
    "Column 18": "value 4",
    "Column 19": "value 4",
    "Column 20": "value 4",
    "Column 21": "value 4",
    "Column 22": "value 4",
  },
  {
    "Column 1": "value 5",
    "Column 2": "value 5",
    "Column 3": "value 5",
    "Column 4": "value 5",
    "Column 5": "value 5",
    "Column 6": "value 5",
    "Column 7": "value 5",
    "Column 8": "value 5",
    "Column 9": "value 5",
    "Column 10": "value 5",
    "Column 11": "value 5",
    "Column 12": "value 5",
    "Column 13": "value 5",
    "Column 14": "value 5",
    "Column 15": "value 5",
    "Column 16": "value 5",
    "Column 17": "value 5",
    "Column 18": "value 5",
    "Column 19": "value 5",
    "Column 20": "value 5",
    "Column 21": "value 5",
    "Column 22": "value 5",
  },
];

function App() {
  const handleReport = () => {
    const wb = XLSX.utils.book_new();
    data.unshift(
      {
        "Column 1": "",
        "Column 2": "",
        "Column 3": "",
        "Column 4": "",
        "Column 5": "",
        "Column 6": "",
        "Column 7": "",
        "Column 8": "",
        "Column 9": "",
        "Column 10": "",
        "Column 11": "",
        "Column 12": "",
        "Column 13": "",
        "Column 14": "",
        "Column 15": "",
        "Column 16": "",
        "Column 17": "",
        "Column 18": "",
        "Column 19": "",
        "Column 20": "",
        "Column 21": "",
        "Column 22": "",
      },
      {
        "Column 1": "",
        "Column 2": "",
        "Column 3": "",
        "Column 4": "",
        "Column 5": "",
        "Column 6": "",
        "Column 7": "",
        "Column 8": "",
        "Column 9": "",
        "Column 10": "",
        "Column 11": "",
        "Column 12": "",
        "Column 13": "",
        "Column 14": "",
        "Column 15": "",
        "Column 16": "",
        "Column 17": "",
        "Column 18": "",
        "Column 19": "",
        "Column 20": "",
        "Column 21": "",
        "Column 22": "",
      },
      {
        "Column 1": "",
        "Column 2": "",
        "Column 3": "",
        "Column 4": "",
        "Column 5": "",
        "Column 6": "",
        "Column 7": "",
        "Column 8": "",
        "Column 9": "",
        "Column 10": "",
        "Column 11": "",
        "Column 12": "",
        "Column 13": "",
        "Column 14": "",
        "Column 15": "",
        "Column 16": "",
        "Column 17": "",
        "Column 18": "",
        "Column 19": "",
        "Column 20": "",
        "Column 21": "",
        "Column 22": "",
      }
    );
    const ws = XLSX.utils.json_to_sheet(data, { skipHeader: true });
    ws.A1 = { t: "s", v: "" };
    ws.A2 = { t: "s", v: "" };
    ws.A3 = { t: "s", v: "Column 1" };
    ws.B3 = { t: "s", v: "Column 2" };
    ws.C3 = { t: "s", v: "Column 3" };
    ws.D3 = { t: "s", v: "Column 4" };
    ws.E3 = { t: "s", v: "Column 5" };
    ws.F3 = { t: "s", v: "Column 6" };
    ws.G3 = { t: "s", v: "Column 7" };
    ws.H3 = { t: "s", v: "Column 8" };
    ws.I3 = { t: "s", v: "Column 9" };
    ws.J3 = { t: "s", v: "Column 10" };
    ws.K1 = { t: "s", v: "Header 1" };
    ws.K2 = { t: "s", v: "Header 3" };
    ws.N2 = { t: "s", v: "Header 4" };
    ws.K3 = { t: "s", v: "Column 11" };
    ws.L3 = { t: "s", v: "Column 12" };
    ws.M3 = { t: "s", v: "Column 13" };
    ws.N3 = { t: "s", v: "Column 14" };
    ws.O3 = { t: "s", v: "Column 15" };
    ws.P3 = { t: "s", v: "Column 16" };
    ws.Q1 = { t: "s", v: "Header 2" };
    ws.Q2 = { t: "s", v: "Header 3" };
    ws.T2 = { t: "s", v: "Header 4" };
    ws.Q3 = { t: "s", v: "Column 17" };
    ws.R3 = { t: "s", v: "Column 18" };
    ws.S3 = { t: "s", v: "Column 19" };
    ws.T3 = { t: "s", v: "Column 20" };
    ws.U3 = { t: "s", v: "Column 21" };
    ws.V3 = { t: "s", v: "Column 22" };

    const merge = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } },
      { s: { r: 0, c: 10 }, e: { r: 0, c: 15 } },
      { s: { r: 0, c: 16 }, e: { r: 0, c: 21 } },
      { s: { r: 1, c: 10 }, e: { r: 1, c: 12 } },
      { s: { r: 1, c: 13 }, e: { r: 1, c: 15 } },
      { s: { r: 1, c: 16 }, e: { r: 1, c: 18 } },
      { s: { r: 1, c: 19 }, e: { r: 1, c: 21 } },
    ];
    ws["!merges"] = merge;

    const wscols = [
      { wch: 10 },
      { wch: 35 },
      { wch: 15 },
      { wch: 20 },
      { wch: 35 },
      { wch: 20 },
      { wch: 25 },
      { wch: 20 },
      { wch: 20 },
      { wch: 20 },
      { wch: 25 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
    ];
    ws["!cols"] = wscols;

    XLSX.utils.book_append_sheet(wb, ws, "");
    XLSX.writeFile(wb, "download.xlsx");
  };

  return (
    <div className="p-24">
      {/* <button className="block px-5 py-3 text-center font-medium text-white bg-indigo-900 hover:bg-gray-800 rounded" onClick={handleExport}>
        Export Excel
      </button> */}
      <button className="my-3 block px-5 py-3 text-center font-medium text-white bg-indigo-900 hover:bg-gray-800 rounded" onClick={handleReport}>
        Report Excel
      </button>
    </div>
  );
}

export default App;
