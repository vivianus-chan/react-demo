import GC from "@grapecity/spread-sheets";
import { IEventTypeObj } from "@grapecity/spread-sheets-react";
import { Button, Space, Tabs } from "antd";
import "assets/css/spread.scss";
import { ExcelSheet } from "components/ExcelSheet";
import { spreadData } from "data/spread";
import { FC, useEffect, useState } from "react";

const tabs = [
  {
    name: "页签 1",
  },
  {
    name: "页签 2",
  },
];

const SpreadDemo: FC = () => {
  const [spread, setSpread] = useState<GC.Spread.Sheets.Workbook>();
  const [sheet, setSheet] = useState<GC.Spread.Sheets.Worksheet>();
  const [data, setData] = useState<any>();
  const [activeIndex, setActiveIndex] = useState("1");

  useEffect(() => {
    setData(JSON.parse(JSON.stringify(spreadData)));
  }, []);

  const get = () => {
    console.log(data, spreadData);
    console.log(
      `总行数：：${sheet?.getRowCount()}，总列数：：${sheet?.getColumnCount()}`
    );
  };

  const getSel = () => {
    console.log(
      `row: ${sheet?.getActiveRowIndex()}，column：${sheet?.getActiveColumnIndex()}`
    );
  };

  const addRowFromActive = () => {
    const mergeColumnIndex = 1; // 需要合并的列下标
    const activeRowIndex = sheet?.getActiveRowIndex() ?? 0;
    const activeColumnIndex = sheet?.getActiveColumnIndex() ?? 0;
    const activeItem = sheet?.getDataItem(activeRowIndex);
    // const mergeCellItems =
    //   sheet?.getSpans(
    //     new GC.Spread.Sheets.Range(activeRowIndex, mergeColumnIndex, 1, 1)
    //   ) ?? [];
    // const mergeCellItem = mergeCellItems?.[0] ?? {
    //   rowCount: 1,
    //   row: activeRowIndex,
    // };

    //@ts-ignore
    const mergeCellItem = sheet?.getSpan(activeRowIndex, mergeColumnIndex) ?? {
      rowCount: 1,
      row: activeRowIndex,
    };

    console.log(
      "新增并合并单元格：：",
      activeRowIndex,
      activeColumnIndex,
      // mergeCellItems,
      mergeCellItem
    );

    spread?.suspendPaint();
    sheet?.addRows(activeRowIndex + 1, 1);
    sheet?.setArray(activeRowIndex + 1, 0, [
      [null, activeItem.Category, null, null],
    ]);
    sheet?.addSpan(
      mergeCellItem?.row,
      mergeColumnIndex,
      mergeCellItem?.rowCount + 1,
      1
    );
    // .hAlign(GC.Spread.Sheets.HorizontalAlign.center);
    sheet
      ?.getCell(mergeCellItem.row, mergeCellItem.col)
      .vAlign(GC.Spread.Sheets.VerticalAlign.center);
    spread?.resumePaint();
  };

  const addRowFromTail = () => {
    sheet?.addRows(sheet.getRowCount(), 1);
  };

  const addColumnFromTail = () => {
    sheet?.addColumns(sheet.getColumnCount(), 1);
  };

  const deleteRow = () => {
    const activeRowIndex = sheet?.getActiveRowIndex() ?? 0;
    sheet?.deleteRows(activeRowIndex, 1);
  };

  const deleteColumn = () => {
    const activeColumnIndex = sheet?.getActiveColumnIndex() ?? 0;
    sheet?.deleteColumns(activeColumnIndex, 1);
  };

  const setSel = () => {
    sheet?.setActiveCell(1, 3);
  };

  const refresh = () => {
    console.log("刷新啦");
    const count1 = sheet?.getColumnCount() ?? 0; // reset前的列数
    const ColumnWidth = [];
    spread?.suspendPaint();
    for (let i = 0; i < count1; i++) {
      ColumnWidth.push(sheet?.getColumnWidth(i));
    }
    sheet?.reset();
    setData(JSON.parse(JSON.stringify(spreadData)));
    const count2 = sheet?.getColumnCount() ?? 0; // 设置为初始值后，获取最新的列数不对
    console.log(count2, ColumnWidth);
    for (let i = 0; i < ColumnWidth.length; i++) {
      console.log(i, ColumnWidth[i]);
      sheet?.setColumnWidth(i, ColumnWidth[i]);
    }
    spread?.refresh();
    spread?.resumePaint();
  };

  const merge = () => {
    console.log(data, spread);
    spread?.suspendPaint();
    let repeatCount = 1;
    data.forEach((x: any, index: number) => {
      console.log(555, x);
      if (x.Category) {
        if (x.Category === data[index + 1]?.Category) {
          repeatCount++;
        } else {
          sheet?.addSpan(Math.abs(repeatCount - index - 1), 1, repeatCount, 1);
          repeatCount = 1;
        }
      }
    });
    spread?.resumePaint();
  };

  const isCellinSpan = (row: number, col: number) => {
    var ranges =
      sheet?.getSpans(new GC.Spread.Sheets.Range(row, col, 1, 1)) ?? [];
    console.log(`${row}行和${col}列合并单元格区域：：`, ranges);
    if (ranges.length) {
      return true;
    }
    return false;
  };

  const valueChanged = (
    type: IEventTypeObj,
    args: GC.Spread.Sheets.IValueChangedEventArgs
  ) => {
    console.log(type, args);
    const { row, col } = args;
    console.log(
      `valueChanged中判断${row}行和${col}列是否在合并单元格内：：`,
      isCellinSpan(row, col)
    );
  };

  return (
    <Tabs
      className="tab-content"
      defaultActiveKey={activeIndex}
      onChange={(key) => setActiveIndex(key)}
      type="card"
      animated
    >
      {tabs.map((x, index) => (
        <Tabs.TabPane tab={x.name} key={index + 1} className="height100">
          <Space style={{ margin: "10px" }}>
            <Button onClick={() => get()}>获取数据源</Button>
            <Button onClick={() => getSel()}>获取选中</Button>
            <Button onClick={() => setSel()}>选中cell</Button>
            <Button onClick={() => merge()}>合并行</Button>
            <Button onClick={() => refresh()}>刷新</Button>
          </Space>
          <Space style={{ margin: "10px" }}>
            <Button onClick={() => addRowFromActive()}>
              添加行并合并单元格
            </Button>
            <Button onClick={() => addRowFromTail()}>末尾添加行</Button>
            <Button onClick={() => addColumnFromTail()}>末尾添加列</Button>
            <Button onClick={() => deleteRow()}>删除行</Button>
            <Button onClick={() => deleteColumn()}>删除列</Button>
            <Button onClick={() => isCellinSpan(0, 1)}>
              判断是否在合并单元格中
            </Button>
          </Space>

          <ExcelSheet
            data={data}
            spreadSheets={{
              workbookInitialized: (spread) => {
                setSpread(spread);
                setSheet(spread?.getActiveSheet());
              },
              valueChanged,
            }}
          ></ExcelSheet>
        </Tabs.TabPane>
      ))}
    </Tabs>
  );
};

export default SpreadDemo;
