import GC from "@grapecity/spread-sheets";
import {
  Column,
  SpreadSheets,
  Worksheet
} from "@grapecity/spread-sheets-react";
import { useEffect, useState } from "react";

export const ExcelSheet: React.FC<IExcelSheetProps> = (props) => {
  const [spread, setSpread] = useState(null);
  const [sheetName, setSheetName] = useState("假装数据");
  const [hostStyle, setHostStyle] = useState<any>({
    width: "80%",
    height: "80%",
    margin: "30px auto",
  });

  useEffect(() => {
    console.log(props, spread);
  }, []);

  const initSpread = (spread: any) => {
    setSpread(spread);
    spread.suspendPaint();
    let sheet = spread.getActiveSheet();
    // sheet.setRowHeight(col, 147);
    // sheet.setColumnWidth(col, 147);
    // console.log(spread.fromJSON(json))
    var spreadNS = GC.Spread.Sheets;
    // spread.setSheetCount(3);
    // var self = this;
    // spread.bind(spreadNS.Events.ActiveSheetChanged, function (e, args) {
    //   var index = spread.getActiveSheetIndex();
    //   self.activeSheetIndex = index;
    // });
    //   spread.bind(spreadNS.Events.CellClick, function (e, args) {
    //     let sheetArea = args.sheetArea === 0 ? 'sheetCorner' : args.sheetArea === 1 ? 'columnHeader' : args.sheetArea === 2 ? 'rowHeader' : 'viewPort';
    //     let log =
    //         'SpreadEvent: ' + GC.Spread.Sheets.Events.CellClick + ' event called' + '\n' +
    //         'sheetArea: ' + sheetArea + '\n' +
    //         'row: ' + args.row + '\n' +
    //         'col: ' + args.col;
    //     self.setState({ eventLog: log });
    // });
    //   spread.bind(spreadNS.Events.EditStarting, function (e, args) {
    //     let log =
    //         'SpreadEvent: ' + GC.Spread.Sheets.Events.EditStarting + ' event called' + '\n' +
    //         'row: ' + args.row + '\n' +
    //         'column: ' + args.col;
    //     self.setState({ eventLog: log });
    // });
    // spread.bind(spreadNS.Events.EditEnded, function (e, args) {
    //     let log =
    //         'SpreadEvent: ' + GC.Spread.Sheets.Events.EditEnded + ' event called' + '\n' +
    //         'row: ' + args.row + '\n' +
    //         'column: ' + args.col + '\n' +
    //         'text: ' + args.editingText;
    //     self.setState({ eventLog: log });
    // });
    spread.resumePaint();
  };

  return (
    <SpreadSheets
      backColor="#fff"
      // grayAreaBackColor="#E4E4E4"
      hostStyle={hostStyle}
      newTabVisible={false}
      tabStripVisible={true}
      scrollbarMaxAlign={true}
      workbookInitialized={(spread) => initSpread(spread)}
    >
      <Worksheet
        name={sheetName}
        dataSource={props.data}
        autoGenerateColumns={false}
        // frozenRowCount={1}
        // frozenColumnCount={1}
      >
        <Column dataField="Name" width={300}></Column>
        <Column dataField="Category" width={100}></Column>
        <Column dataField="Price" width={100} formatter="$#.00"></Column>
        <Column dataField="Shopping Place" width={100}></Column>
      </Worksheet>
    </SpreadSheets>
  );
};

interface IExcelSheetProps {
  data: any[];
}
