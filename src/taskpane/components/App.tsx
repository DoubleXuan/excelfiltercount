import * as React from "react";
import FilterWithCount from "./FilterWithCount";
import { ThemeProvider } from "@fluentui/react";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App: React.FunctionComponent<AppProps> = (props) => {
  const { isOfficeInitialized } = props;

  if (!isOfficeInitialized) {
    return (
      <div style={{ padding: 20, fontSize: '12px', color: '#666' }}>
        <p>正在加载 Office 环境...</p>
      </div>
    );
  }

  // 这里的 return 里面只有 ThemeProvider 包裹着 FilterWithCount
  // 绝对没有 props.title 的渲染代码
  return (
    <ThemeProvider>
      <FilterWithCount />
    </ThemeProvider>
  );
};

export default App;