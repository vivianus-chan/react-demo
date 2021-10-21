// import GC from '@grapecity/spread-sheets';
import {
    Column,
    SpreadSheets,
    Worksheet
} from "@grapecity/spread-sheets-react";
import { useState } from "react";
import { spreadDate } from "../data/spread";
// GC.Spread.Sheets.LicenseKey = 'sds';

function SpreadDemo() {
  const [spreadBackColor, setSpreadBackColor] = useState("aliceblue");
  const [sheetName, setSheetName] = useState("Goods List");
  const [hostStyle, setHostStyle] = useState<any>({
    width: "800px",
    height: "600px",
  });
  const [columnWidth, setColumnWidth] = useState(100);
  const [data, setData] = useState(spreadDate);

  return (
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
  );
}

export default SpreadDemo;
