import GC from "@grapecity/spread-sheets";
import {
  Column,
  IEventTypeObj,
  SpreadSheets,
  SpreadSheetsProp,
  Worksheet,
  WorksheetProp
} from "@grapecity/spread-sheets-react";
import { Button, Space } from "antd";
import "assets/css/spread.scss";
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

  useEffect(() => {
    console.log(props);
  }, []);

  useEffect(() => {
    sheet?.setDataSource(JSON.parse(JSON.stringify(spreadData)));
  }, [sheet]);

  const get = () => {
    console.log(sheet?.getDataSource(), spreadData);
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

  const refresh = () => {
    const count1 = sheet?.getColumnCount() ?? 0; // reset前的列数
    const columnWidth = [];
    spread?.suspendPaint();
    for (let i = 0; i < count1; i++) {
      columnWidth.push(sheet?.getColumnWidth(i));
    }
    sheet?.reset();
    sheet?.setDataSource(JSON.parse(JSON.stringify(spreadData)));
    const count2 = sheet?.getColumnCount() ?? 0; // 设置为初始值后，获取最新的列数不对
    console.log(`新列数：${count2}`, columnWidth);
    for (let i = 0; i < count2; i++) {
      sheet?.setColumnWidth(i, columnWidth[i]);
    }
    spread?.resumePaint();
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

  const isCellinSpan = (row: number, col: number) => {
    console.log(sheet, spread, spread?.getActiveSheet());
    var ranges =
      sheet?.getSpans(new GC.Spread.Sheets.Range(row, col, 1, 1)) ?? [];
    console.log(`(${row},${col})合并单元格区域：：`, ranges);
    if (ranges.length) {
      return true;
    }
    return false;
  };

  const valueChanged = (
    type: IEventTypeObj,
    args: GC.Spread.Sheets.IValueChangedEventArgs
  ) => {
    console.log(type, args);
    const { row, col } = args;
    console.log(
      `valueChanged中判断(${row},${col})是否在合并单元格内：：`,
      isCellinSpan(row, col)
    );
  };

  const workbookInitialized = (spread: GC.Spread.Sheets.Workbook) => {
    console.log('初始化')
    setSpread(spread);
    setSheet(spread?.getActiveSheet());
  };

  return (
    <>
      <Space style={{ margin: "10px" }}>
        <Button onClick={() => get()}>获取数据源</Button>
        <Button onClick={() => addRowFromTail()}>末尾添加行</Button>
        <Button onClick={() => addColumnFromTail()}>末尾添加列</Button>
        <Button onClick={() => deleteRow()}>删除行</Button>
        <Button onClick={() => deleteColumn()}>删除列</Button>
        <Button onClick={() => getSel()}>获取选中</Button>
        <Button onClick={() => setSel()}>选中cell</Button>
        <Button onClick={() => refresh()}>刷新</Button>
      </Space>
      <Space style={{ margin: "10px" }}>
        <Button onClick={() => addRowFromActive()}>添加行并合并单元格</Button>
        <Button onClick={() => merge()}>合并行</Button>
        <Button onClick={() => isCellinSpan(0, 1)}>
          判断是否在合并单元格中
        </Button>
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
    </>
  );
};

interface IExcelSheetProps {
  sheetName?: string;
  spreadSheets?: SpreadSheetsProp;
  worksheet?: WorksheetProp;
}
