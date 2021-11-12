import GC from "@grapecity/spread-sheets";
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
  const [sheet, setSheet] = useState<GC.Spread.Sheets.Worksheet>();
  const [spread, setSpread] = useState<GC.Spread.Sheets.Workbook>();
  const [data, setData] = useState<any>();
  const [activeIndex, setActiveIndex] = useState("1");

  useEffect(() => {
    setData(spreadData);
  }, []);

  const get = () => {
    console.log(data);
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
    const activeRowIndex = sheet?.getActiveRowIndex() ?? 0;
    const activeItem = sheet?.getDataItem(activeRowIndex);
    console.log(activeItem);
    spread?.suspendPaint();
    sheet?.addRows(activeRowIndex + 1, 1);
    sheet?.setArray(activeRowIndex + 1, 0, [
      [null, activeItem.Category, null, null],
    ]);
    sheet?.addSpan(
      activeRowIndex,
      1,
      2,
      1,
      GC.Spread.Sheets.SheetArea.viewport
    );
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
    const activeRowIndex = sheet?.getActiveRowIndex() ?? 0;
    sheet?.deleteColumns(activeRowIndex, 1);
  };

  const setSel = () => {
    sheet?.setActiveCell(1, 3);
  };

  const refresh = () => {
    // sheet?.reset()
    // console.log("刷新啦");
    // spread?.suspendPaint();
    // setData(spreadData);
    spread?.refresh();
    // spread?.resumePaint();
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
            <Button onClick={() => addRowFromActive()}>添加行</Button>
            <Button onClick={() => addRowFromTail()}>末尾添加行</Button>
            <Button onClick={() => addColumnFromTail()}>末尾添加列</Button>
            <Button onClick={() => deleteRow()}>删除行</Button>
            <Button onClick={() => deleteColumn()}>删除列</Button>
          </Space>

          <ExcelSheet
            data={data}
            bindSpread={(spread) => {
              setSpread(spread);
              setSheet(spread?.getActiveSheet());
            }}
          ></ExcelSheet>
        </Tabs.TabPane>
      ))}
    </Tabs>
  );
};

export default SpreadDemo;
