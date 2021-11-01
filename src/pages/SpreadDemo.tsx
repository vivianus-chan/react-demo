import { Tabs } from "antd";
import "assets/css/spread.scss";
import { ExcelSheet } from "components/ExcelSheet";
import { FC, useEffect, useState } from "react";
import { spreadData } from "../data/spread";

const tabs = [
  {
    name: "页签 1",
  },
  {
    name: "页签 2",
  },
];

const SpreadDemo: FC = () => {
  const [data, setData] = useState<any>();
  const [activeIndex, setActiveIndex] = useState("1");

  useEffect(() => {
    setData(spreadData);
  }, []);

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
          <ExcelSheet data={data}></ExcelSheet>
        </Tabs.TabPane>
      ))}
    </Tabs>
  );
};

export default SpreadDemo;
