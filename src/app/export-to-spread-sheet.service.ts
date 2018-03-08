import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { WorkBook } from 'xlsx';
import { SpreadsheetTypes } from './spreadsheet-types.enum';
import { Column } from './column';
@Injectable()
export class ExportToSpreadSheetService {

  constructor() {

  }
  // takes any Object tries to put it into a spreadsheet
  public generateWorkSheet(
    title: string,
    fileName: string,
    columns: Column[],
    data: Object[],
    fileType: SpreadsheetTypes,
    optionalHeaders?: string[],
    optionalFooters?: string[]): void {

    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet([[title, 'generated on', new Date(Date.now()).toDateString()]]);

    const columnNames = [columns.map(col => col.HeaderName ? col.HeaderName : col.FieldName)];
    const fieldNames =  columns.map(col => col.FieldName);

    // Make sure data is in right order
    // O(n^3)
    const mappedData = data.map(dataRow => {
      const rowData = [];
      const objectProps = Object.getOwnPropertyNames(dataRow);
      for (const fieldName of fieldNames) {

      objectProps.forEach(prop => {
          if (fieldName === prop) {
            rowData.push(dataRow[prop]);
          }
        });

      }
      return rowData;

    });

    console.log(mappedData);


    if (optionalHeaders != null) {

      XLSX.utils.sheet_add_aoa(ws, [optionalHeaders], { origin: { r: 1, c: 0 } });
      XLSX.utils.sheet_add_aoa(ws, columnNames, { origin: { r: 2, c: 0 } });
      XLSX.utils.sheet_add_aoa(ws, mappedData, { origin: { r: 3, c: 0 } });

    } else {
      XLSX.utils.sheet_add_aoa(ws, columnNames, { origin: { r: 1, c: 0 } });

      XLSX.utils.sheet_add_json(ws, mappedData, { origin: { r: 2, c: 0 } });

    }
    if (optionalFooters) {
      XLSX.utils.sheet_add_aoa(ws, [optionalFooters], { origin: { r: optionalHeaders != null ? data.length + 3 : data.length + 2, c: 0 } });
    }





    const wb: XLSX.WorkBook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, fileName);
    XLSX.writeFile(wb, fileName + '.' + fileType);



  }





}
