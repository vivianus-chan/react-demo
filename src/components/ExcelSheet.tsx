import {
  Column,
  SpreadSheets,
  SpreadSheetsProp,
  Worksheet,
  WorksheetProp
} from "@grapecity/spread-sheets-react";
import { useEffect, useState } from "react";

export const ExcelSheet: React.FC<IExcelSheetProps> = (props) => {
  const [hostStyle, setHostStyle] = useState<any>({
    width: "80%",
    height: "50%",
    margin: "30px auto",
  });

  useEffect(() => {
    console.log(props);
  }, []);

  return (
    <SpreadSheets
      backColor="#fff"
      // grayAreaBackColor="#E4E4E4"
      hostStyle={hostStyle}
      newTabVisible={false}
      tabStripVisible={true}
      scrollbarMaxAlign={true}
      {...props.spreadSheets}
    >
      <Worksheet
        name={props.sheetName}
        autoGenerateColumns={false}
        // selectionUnit={GC.Spread.Sheets.SelectionUnit.row}
        // selectionBorderColor="red"
        // selectionBackColor="transparent"
        // frozenRowCount={1}
        // frozenColumnCount={1}
        // frozenlineColor="Transparent"
        {...props.worksheet}
      >
        <Column dataField="Name" width={300}></Column>
        <Column dataField="Category" width={100}></Column>
        <Column dataField="Price" width={110} formatter="$#.00"></Column>
        <Column dataField="Shopping Place" width={120}></Column>
      </Worksheet>
    </SpreadSheets>
  );
};

interface IExcelSheetProps {
  sheetName?: string;
  spreadSheets?: SpreadSheetsProp;
  worksheet?: WorksheetProp;
}
