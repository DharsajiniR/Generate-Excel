
import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {
  constructor() {
  }
  generateExcel() {
      
    //Create workbook and worksheet
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet('Sample');
   
   
    worksheet.columns = [
      { header: 'itemId', width: 10, key: 'id' },
      { header: 'Name', width: 10, key: 'name' },
      { header: 'Age', width: 10, key: 'age' },
    ];
    
    worksheet.addRow({
      id: 1,
      name: 'dj',
      age: '20'
    })
    worksheet.getRow(1).fill ={
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'a1a1a1' }
        }
    worksheet.getRow(1).font = {  bold: true};
    
    
    worksheet.getColumn(2).width = 30;
    worksheet.getColumn(3).width = 30;
   
    //Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, 'Sample.xlsx');
    })
  }
}