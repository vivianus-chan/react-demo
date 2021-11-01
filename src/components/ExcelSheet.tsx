import {
  Column,
  SpreadSheets,
  Worksheet
} from "@grapecity/spread-sheets-react";
import { spreadData } from "data/spread";
import { FC, useEffect, useState } from "react";
// import GC from '@grapecity/spread-sheets';
// GC.Spread.Sheets.LicenseKey = 'sds';

export const ExcelSheet: FC<IExcelSheetProps> = (props) => {
  const [spreadBackColor, setSpreadBackColor] = useState("#fff");
  const [sheetName, setSheetName] = useState("假装数据");
  const [hostStyle, setHostStyle] = useState<any>({
    width: "100%",
    height: "100%",
  });
  const [columnWidth, setColumnWidth] = useState(100);
  const [data, setData] = useState<any>(spreadData);

  useEffect(() => {
    console.log(props);
  }, [props.data]);

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
};

interface IExcelSheetProps {
  data: any[];
}
