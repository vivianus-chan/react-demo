// import GC from '@grapecity/spread-sheets';
import "@/assets/css/spread.scss";
import {
  Column,
  SpreadSheets,
  Worksheet
} from "@grapecity/spread-sheets-react";
import { Tabs } from "antd";
import { FC, useEffect, useState } from "react";
import { spreadDate } from "../data/spread";

// GC.Spread.Sheets.LicenseKey = 'sds';

const { TabPane } = Tabs;

const SpreadDemo: FC = () => {
  const [spreadBackColor, setSpreadBackColor] = useState("#fff");
  const [sheetName, setSheetName] = useState("假装数据");
  const [hostStyle, setHostStyle] = useState<any>({
    width: "100%",
    height: "100%",
  });
  const [columnWidth, setColumnWidth] = useState(100);
  const [data, setData] = useState<any>(spreadDate);
  const [activeIndex, setActiveIndex] = useState("1");

  useEffect(() => {
    setData(spreadDate);
  });

  return (
    <Tabs
      className="tab-content"
      defaultActiveKey={activeIndex}
      onChange={(key) => setActiveIndex(key)}
      type="card"
      animated
    >
      <TabPane tab="页签 1" key="1" className="height100">
        <SpreadSheets backColor={spreadBackColor} hostStyle={hostStyle}>
          <Worksheet name={sheetName} dataSource={data}>
            <Column dataField="Name" width={300}></Column>
            <Column dataField="Category" width={columnWidth}></Column>
            <Column
              dataField="Price"
              width={columnWidth}
              formatter="$#.00"
            ></Column>
            <Column dataField="Shopping Place" width={columnWidth}></Column>
          </Worksheet>
        </SpreadSheets>
      </TabPane>
      <TabPane tab="页签 2" key="2" className="height100">
        <SpreadSheets backColor={spreadBackColor} hostStyle={hostStyle}>
          <Worksheet name={sheetName} dataSource={data}>
            <Column dataField="Name" width={300}></Column>
            <Column dataField="Category" width={columnWidth}></Column>
          </Worksheet>
        </SpreadSheets>
      </TabPane>
    </Tabs>
  );
};

export default SpreadDemo;
