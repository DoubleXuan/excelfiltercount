import "es6-promise/auto"; 
import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./components/App";
import { initializeIcons } from "@fluentui/react";

/* global document, Office, module, require */

initializeIcons();

// 1. 注意：这里不再定义 const title = "增强筛选工具";
// 我们直接传一个空字符串，或者干脆删掉这个属性（需要在 App.tsx 接口里也去掉，传空串最省事）

const render = (Component) => {
  ReactDOM.render(
    <Component title="" isOfficeInitialized={true} />, // 传空字符串
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}