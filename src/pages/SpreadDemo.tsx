import { Tabs } from "antd";
import "assets/css/spread.scss";
import { ExcelSheet } from "components/ExcelSheet";
import { FC, useState } from "react";

const tabs = [
  {
    name: "页签 1",
    sheetName: "假装数据",
  },
  {
    name: "页签 2",
    sheetName: "东方数据",
  },
];

const SpreadDemo: FC = () => {
  const [activeIndex, setActiveIndex] = useState("1");

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
          <ExcelSheet sheetName={x.sheetName} spreadSheets={{}}></ExcelSheet>
        </Tabs.TabPane>
      ))}
    </Tabs>
  );
};

export default SpreadDemo;
