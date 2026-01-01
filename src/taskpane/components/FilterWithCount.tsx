/* global Excel, console */
import * as React from 'react';
import { DefaultButton, PrimaryButton, Checkbox, Stack, Text, SearchBox, IIconProps } from '@fluentui/react';

interface FilterItem {
  value: string;
  count: number;
  isChecked: boolean;
}

const exportIcon: IIconProps = { iconName: 'Download' };
const filterIcon: IIconProps = { iconName: 'Filter' };

const FilterWithCount: React.FunctionComponent = () => {
  const [items, setItems] = React.useState<FilterItem[]>([]);
  const [columnHeader, setColumnHeader] = React.useState<string>("");
  const [colIndex, setColIndex] = React.useState<number>(-1);
  const [searchText, setSearchText] = React.useState<string>("");

  const filteredItems = React.useMemo(() => {
    if (!searchText.trim()) {
      return items;
    }
    const lowerText = searchText.toLowerCase();
    return items.filter(item => 
      String(item.value).toLowerCase().includes(lowerText)
    );
  }, [items, searchText]);

  const isAllVisibleSelected = filteredItems.length > 0 && filteredItems.every(item => item.isChecked);
  const isIndeterminate = filteredItems.some(item => item.isChecked) && !isAllVisibleSelected;

  // --- 核心功能函数 ---

  const loadColumnData = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["columnIndex"]);
        
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
        
        await context.sync();

        const activeColIndex = range.columnIndex;
        const startColIndex = usedRange.columnIndex;
        
        if (activeColIndex < startColIndex || activeColIndex >= startColIndex + usedRange.columnCount) {
          console.error("请选择数据区域内的单元格");
          return;
        }

        const relativeCol = activeColIndex - startColIndex;
        const allValues = usedRange.values;
        
        const header = String(allValues[0][relativeCol]);
        const dataColumn = [];
        for (let i = 1; i < allValues.length; i++) {
          dataColumn.push(String(allValues[i][relativeCol])); 
        }

        const counts: Record<string, number> = {};
        dataColumn.forEach(val => {
          counts[val] = (counts[val] || 0) + 1;
        });

        const filterItems: FilterItem[] = Object.keys(counts).map(key => ({
          value: key,
          count: counts[key],
          isChecked: true 
        })).sort((a, b) => b.count - a.count);

        setColumnHeader(header);
        setColIndex(activeColIndex);
        setItems(filterItems);
        setSearchText(""); 
      });
    } catch (error) {
      console.error(error);
    }
  };

  const executeFilterInExcel = async (targetValues?: string[]) => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        sheet.load("autoFilter");
        usedRange.load("columnIndex"); 
        
        await context.sync();

        const relativeIndex = colIndex - usedRange.columnIndex;
        if (relativeIndex < 0) return;

        let finalValues: string[] = [];
        if (targetValues) {
            finalValues = targetValues;
        } else {
            finalValues = items.filter(i => i.isChecked).map(i => i.value);
        }

        const isSelectAll = finalValues.length === items.length;

        if (isSelectAll || finalValues.length === 0) {
            if ((sheet.autoFilter as any).removeFilterCriteria) {
               (sheet.autoFilter as any).removeFilterCriteria(relativeIndex);
            } else {
               (sheet.autoFilter as any).apply(usedRange, relativeIndex, {
                 filterOn: Excel.FilterOn.values,
                 values: items.map(i => i.value) 
               });
            }
        } else {
            sheet.autoFilter.apply(usedRange, relativeIndex, {
              filterOn: Excel.FilterOn.values,
              values: finalValues
            });
        }
        await context.sync();
      });
    } catch (error) {
      console.error("筛选出错:", error);
    }
  };

  // --- 导出统计结果 ---
  const handleExport = async () => {
    try {
      await Excel.run(async (context) => {
        const newSheet = context.workbook.worksheets.add();
        
        // 【关键修改】动态设置表头名称
        // 如果 columnHeader 有值，就用它（例如“省份”）；否则兜底显示“内容”
        const headerName = columnHeader || "内容";
        
        // 定义数据数组
        const exportData: (string | number)[][] = [[headerName, "数量"]];
        
        items.forEach(item => {
            exportData.push([item.value, item.count]);
        });

        const range = newSheet.getRangeByIndexes(0, 0, exportData.length, 2);
        range.values = exportData;

        const headerRange = newSheet.getRange("A1:B1");
        headerRange.format.font.bold = true; 
        headerRange.format.fill.color = "#E1DFDD"; 
        
        newSheet.getUsedRange().format.autofitColumns();
        newSheet.activate();

        await context.sync();
      });
    } catch (error) {
      console.error("导出失败:", error);
    }
  };

  // --- 交互处理 ---

  const handleApplyButton = () => {
    if (searchText.trim()) {
        const lowerText = searchText.toLowerCase();
        const activeValues = items
            .filter(i => i.isChecked && String(i.value).toLowerCase().includes(lowerText))
            .map(i => i.value);

        setItems(prev => prev.map(item => ({
            ...item,
            isChecked: activeValues.includes(item.value)
        })));

        executeFilterInExcel(activeValues);
    } else {
        executeFilterInExcel();
    }
  };

  const handleSearchEnter = () => {
    handleApplyButton();
  };

  const handleCheckboxChange = (value: string, checked?: boolean) => {
    setItems(prev => prev.map(item => 
      item.value === value ? { ...item, isChecked: !!checked } : item
    ));
  };

  const handleSelectAll = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
    const visibleValues = new Set(filteredItems.map(i => i.value));
    setItems(prev => prev.map(item => {
      if (visibleValues.has(item.value)) {
        return { ...item, isChecked: !!checked };
      }
      return item;
    }));
  };

  return (
    <Stack tokens={{ childrenGap: 10 }} style={{ padding: 10, height: '100vh', boxSizing: 'border-box' }}>
      <DefaultButton text="读取当前列数据" onClick={loadColumnData} primary />
      
      {columnHeader && (
        <Stack horizontal verticalAlign="baseline" tokens={{ childrenGap: 5 }}>
          <Text variant="large">正在筛选: <b>{columnHeader}</b></Text>
        </Stack>
      )}

      <div style={{ flexGrow: 1, display: 'flex', flexDirection: 'column', border: '1px solid #edebe9', borderRadius: 2, overflow: 'hidden' }}>
        
        {items.length > 0 && (
          <div style={{ background: '#f3f2f1', borderBottom: '1px solid #edebe9', padding: 8 }}>
            <SearchBox 
              placeholder="搜索... (回车应用)" 
              value={searchText}
              onChange={(_, newValue) => setSearchText(newValue || "")}
              onSearch={handleSearchEnter}
              underlined={false}
              styles={{ root: { marginBottom: 8 } }}
            />
            <Checkbox 
              label={`(全选结果) - ${filteredItems.length} 项`}
              checked={isAllVisibleSelected}
              indeterminate={isIndeterminate}
              onChange={handleSelectAll}
            />
          </div>
        )}

        <div style={{ flexGrow: 1, overflowY: 'auto', padding: 10, background: 'white' }}>
          {filteredItems.map((item) => (
            <Checkbox 
              key={item.value}
              label={`${item.value} (${item.count})`} 
              checked={item.isChecked}
              onChange={(_, checked) => handleCheckboxChange(item.value, checked)}
              styles={{ root: { marginBottom: 8 } }}
            />
          ))}
          
          {items.length > 0 && filteredItems.length === 0 && (
            <Text style={{ color: '#666', fontStyle: 'italic' }}>未找到匹配项</Text>
          )}

          {items.length === 0 && (
            <Text style={{ color: '#666' }}>请选择一列并点击读取按钮</Text>
          )}
        </div>
      </div>

      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <DefaultButton 
            text="导出统计" 
            iconProps={exportIcon}
            onClick={handleExport} 
            disabled={items.length === 0} 
            styles={{ root: { flex: 1 } }} 
        />
        <PrimaryButton 
            text="应用筛选" 
            iconProps={filterIcon}
            onClick={handleApplyButton} 
            disabled={colIndex === -1 || items.length === 0} 
            styles={{ root: { flex: 1 } }} 
        />
      </Stack>
    </Stack>
  );
};

export default FilterWithCount;