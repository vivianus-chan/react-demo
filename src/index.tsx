import GC from "@grapecity/spread-sheets";
import "@grapecity/spread-sheets-resources-zh";
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.css";
import React from "react";
import ReactDOM from "react-dom";
import App from "./App";
import "./index.css";
import reportWebVitals from "./reportWebVitals";
// GC.Spread.Sheets.LicenseKey = 'sds';
GC.Spread.Common.CultureManager.culture("zh-cn");

ReactDOM.render(<App />, document.getElementById("root"));

// If you want to start measuring performance in your app, pass a function
// to log results (for example: reportWebVitals(console.log))
// or send to an analytics endpoint. Learn more: https://bit.ly/CRA-vitals
reportWebVitals();
