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
    sheet?.setDataSource(JSON.parse(JSON.stringify(spreadData[sheetName])));
    sheet?.bindColumns(column);
    sheet?.frozenColumnCount(1);
    sheet!.options.frozenlineColor = "Transparent";
    spread?.resumePaint();
  };

  const workbookInitialized = (spread: GC.Spread.Sheets.Workbook) => {
    console.log("workbookInitialized");
    setSpread(spread);
    setSheet(spread?.getActiveSheet());
  };

  useEffect(() => {
    console.log(props);
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
        <Worksheet name={props.sheetName}></Worksheet>
      </SpreadSheets>
    </>
  );
};

interface IExcelSheetProps {
  sheetName: string;
  spreadSheets?: SpreadSheetsProp;
}
