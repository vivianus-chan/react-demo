import GC from "@grapecity/spread-sheets";
import {
  IEventTypeObj,
  SpreadSheets,
  SpreadSheetsProp,
  Worksheet
} from "@grapecity/spread-sheets-react";
import { Button, Space } from "antd";
import "assets/css/spread.scss";
import { column } from "data/column";
import { spreadData } from "data/spread";
import { useEffect, useState } from "react";
const buttonStyle = new GC.Spread.Sheets.CellTypes.Button();
const buttonStyle1 = new GC.Spread.Sheets.Style();
buttonStyle1.cellButtons = [
  {
    caption: "选择",
    useButtonStyle: true,
    visibility: GC.Spread.Sheets.ButtonVisibility.onSelected,
    command: (
      sheet: GC.Spread.Sheets.Worksheet,
      row: number,
      col: number,
      option: any
    ) => {},
  },
];
const defaultStyle = new GC.Spread.Sheets.Style();
// defaultStyle.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
defaultStyle.vAlign = GC.Spread.Sheets.VerticalAlign.center;

export const ExcelSheet: React.FC<IExcelSheetProps> = (props) => {
  const [spread, setSpread] = useState<GC.Spread.Sheets.Workbook>();
  const [sheet, setSheet] = useState<GC.Spread.Sheets.Worksheet>();
  const [hostStyle, setHostStyle] = useState<any>({
    width: "80%",
    height: "50%",
    margin: "30px auto",
  });

  const get = () => {
    console.log(sheet?.getDataSource(), spreadData);
    console.log(
      `总行数：：${sheet?.getRowCount()}，总列数：：${sheet?.getColumnCount()}`
    );
  };

  const getSel = () => {
    const activeRowIndex = sheet?.getActiveRowIndex() ?? 0;
    const activeColumnIndex = sheet?.getActiveColumnIndex() ?? 0;
    isCellinSpan(activeRowIndex, activeColumnIndex);
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
    const rowCount = sheet?.getRowCount() ?? 0;
    sheet?.addRows(rowCount, 1);
    sheet?.setFormula(
      rowCount,
      2,
      `=SUM(C1:C${rowCount})`,
      GC.Spread.Sheets.SheetArea.viewport
    );
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

  const merge = () => {
    spread?.suspendPaint();
    let repeatCount = 1;
    const data: any = sheet?.getDataSource() ?? [];
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

  const copy = () => {
    const activeRowIndex = sheet?.getActiveRowIndex() ?? 0;
    const activeColumnIndex = sheet?.getActiveColumnIndex() ?? 0;
    // sheet?.addRows(activeRowIndex + 1, 6);
    sheet?.copyTo(
      activeRowIndex,
      -1,
      activeRowIndex + 2,
      -1,
      2,
      -1,
      GC.Spread.Sheets.CopyToOptions.all
    );
  };

  const collapsed = () => {
    const CollapsedArr = sheet?.outlineColumn.getCollapsed();
    CollapsedArr.forEach((x: boolean, i: number) => {
      sheet?.outlineColumn.setCollapsed(i, false);
    });
  };

  const isCellinSpan = (
    row: number,
    col: number,
    sht?: GC.Spread.Sheets.Worksheet
  ) => {
    console.log(sheet, spread);
    const sheet2 = sheet ?? sht;
    var ranges =
      sheet2?.getSpans(new GC.Spread.Sheets.Range(row, col, 1, 1)) ?? [];
    console.log(`(${row},${col})合并单元格区域：：`, ranges);
    if (ranges.length) {
      return ranges;
    }
    return false;
  };

  const valueChanged = (
    type: IEventTypeObj,
    args: GC.Spread.Sheets.IValueChangedEventArgs
  ) => {
    // console.log(type, args);
    const { row, col, sheet, newValue } = args;
    const spanArea = isCellinSpan(row, col, sheet);
    if (spanArea) {
      for (
        let i = spanArea[0].row;
        i < spanArea[0].row + spanArea[0].rowCount;
        i++
      ) {
        for (
          let j = spanArea[0].col;
          j < spanArea[0].col + spanArea[0].colCount;
          j++
        ) {
          sheet.setValue(i, j, newValue);
        }
      }
    }
  };

  const refresh = () => {
    sheet?.reset();
    initSpread();
  };

  const initSpread = () => {
    spread?.suspendPaint();
    const { sheetName } = props;
    column.forEach((x) => {
      x.displayName = x.displayName.replace(/（.*?）/g, `（${sheetName}）`);
    });
    console.log(column);
    const data: any[] = spreadData[sheetName];
    for (var r = 0; r < data.length; r++) {
      sheet?.getCell(r, 0).textIndent(data[r].level);
    }
    sheet?.setDataSource(JSON.parse(JSON.stringify(data)));
    sheet?.bindColumns(column);
    initOutlineColumn();
    sheet?.frozenColumnCount(1);
    // for (var r = 0; r < data.length; r++) {
    //   if (data[r].level === 2) {
    //     sheet?.outlineColumn.setCollapsed(r, true);
    //   }
    // }
    sheet!.options.frozenlineColor = "Transparent";
    // spread!.options.allowUserEditFormula = false;
    sheet?.setStyle(0, 0, buttonStyle1);
    sheet?.setRowHeight(0, 30, GC.Spread.Sheets.SheetArea.colHeader);
    //set default row height and column width
    sheet!.defaults.rowHeight = 45;
    // sheet!.defaults.colWidth = 150;
    sheet?.setDefaultStyle(defaultStyle, GC.Spread.Sheets.SheetArea.viewport);
    spread?.resumePaint();
  };

  const initOutlineColumn = () => {
    const { sheetName } = props;
    const data: any[] = spreadData[sheetName];
    sheet?.outlineColumn.options({
      columnIndex: 0,
      showImage: false,
      showCheckBox: false,
      maxLevel: 10,
    });
    sheet?.showRowOutline(false);
    sheet?.outlineColumn.refresh();
  };

  const workbookInitialized = (spread: GC.Spread.Sheets.Workbook) => {
    console.log("workbookInitialized");
    setSpread(spread);
    setSheet(spread?.getActiveSheet());
  };

  const splitSum = (cells: string, count = 3) => {
    const arr = cells.split(",");
    const res: any = [];
    for (let i = 0; i < arr.length; i += count) {
      res.push(arr.slice(i, i + count));
    }
    const sumStr = res?.map((cur: string) => `SUM(${cur})`).join("+");
    console.log(arr, res, sumStr);
    return sumStr;
  };

  useEffect(() => {
    console.log(props);
    splitSum("1,2,3,4,5,6,7,8,9,10");
  }, []);

  useEffect(() => {
    sheet && initSpread();
  }, [sheet]);

  return (
    <>
      <Space style={{ margin: "10px" }}>
        <Button onClick={() => get()}>获取数据源</Button>
        <Button onClick={() => addRowFromTail()}>末尾添加行</Button>
        <Button onClick={() => addColumnFromTail()}>末尾添加列</Button>
        <Button onClick={() => deleteRow()}>删除行</Button>
        <Button onClick={() => deleteColumn()}>删除列</Button>
        <Button onClick={() => getSel()}>
          获取选中并判断是否在合并单元格内
        </Button>
        <Button onClick={() => setSel()}>选中cell</Button>
        <Button onClick={() => refresh()}>刷新</Button>
      </Space>
      <Space style={{ margin: "10px" }}>
        <Button onClick={() => addRowFromActive()}>添加行并合并单元格</Button>
        <Button onClick={() => merge()}>合并行</Button>
        <Button onClick={() => copy()}>复制行</Button>
        <Button onClick={() => collapsed()}>展开行</Button>
      </Space>
      <SpreadSheets
        backColor="#fff"
        // grayAreaBackColor="#E4E4E4"
        hostStyle={hostStyle}
        newTabVisible={false}
        tabStripVisible={true}
        scrollbarMaxAlign={true}
        workbookInitialized={workbookInitialized}
        valueChanged={valueChanged}
        {...props.spreadSheets}
      >
        <Worksheet name={props.sheetName} isProtected={false}></Worksheet>
      </SpreadSheets>
    </>
  );
};

interface IExcelSheetProps {
  sheetName: string;
  spreadSheets?: SpreadSheetsProp;
}
