import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { ExportToSpreadSheetService } from './export-to-spread-sheet.service';
import { Column } from './column';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  constructor(private ExportToSpreadSheet: ExportToSpreadSheetService) {

  }
  title = 'TEST DATA';
  data =
  [
    {
    Name: 'name',
    SecondName: 'surname'
   },
   {
    Name: 'nam2',
    SecondName: 'surn2ame'
   },
   {
    Name: 'nam3e',
    SecondName: 'surn3ame'
   },
   {
    Name: 'nam4e',
    SecondName: 'su4name'
   },
   {
    Name: 'name5',
    SecondName: 'surn5ame'
   },
   {
    SecondName: 'sur6name',
    Name: 'nam6e'
   }

  ];


  cols: Column[] =
  [

    {FieldName: 'SecondName', HeaderName: 'Second Name'},
    {FieldName: 'Name', HeaderName: 'First Name'}
  ];

  export(event) {
    this.ExportToSpreadSheet
    .generateWorkSheet('Example', 'example', this.cols, this.data, event,
    ['HELLO WORLD', 'how are you'], ['length:', this.data.length.toString()]);

  }


}
