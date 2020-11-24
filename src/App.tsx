import React from 'react';
import logo from './logo.svg';
import './App.css';
import { ExcelService } from './services/excelService';
import reportWebVitals from './reportWebVitals';

function App() {

  const  generateExcel = (e: any) => {
    const rows = [];
    const styles  = [];
    const fileName = 'export.xlsx';
    const headers = ['Name', 'ID'];
    const lastCol = headers.length - 1;
    rows.push([], headers);
    styles.push(
      // styles can be applied to row, col or cell, but if you do it more then once, only the latest will apply
      {col: '0-' + lastCol, width: 20},
      {row: rows.length - 1, height: 20},
      {row: rows.length - 1, col: '0-' + lastCol, font: 'Arial 10 #000000 B', wrapText: 1,},
    );
    rows.push(['Surya', '777']);
    const excelObj = new ExcelService();
    excelObj.generate(fileName, [{sheetTitle: 'Sheet 1', rows: rows, styles: styles}])

  }

  const  generateReportTemplateExcel = (e: any) => {
    const excelObj = new ExcelService();
    const rows = [];
    const styles  = [];
    const fileName = 'export.xlsx';
    const reportName = 'Active Assets Report';
    const headers = ['Name', 'ID', 'Location', 'Category'];
    const lastCol = headers.length - 1;

    rows.push([], ['REPORT NAME', reportName]);
    rows[0][lastCol - 1] = 'Time Stamp';
    rows[0][lastCol] = excelObj.toExcelLocalTime(new Date());
    rows.push([], ['Location', 'Memphis']);

    rows.push([], ['Report Type', 'Scheduled']);

    rows.push([], ['Status', 'InService']);

    rows.push(['INTERVAL', 'Shiftwise(Shift 1, Shift 2)']);
    rows.push(['FROM', 'July 4 , 2020']);
    rows.push(['TO', 'Jan 1, 2020']);
    rows.push(['FREQUENCY', '1']);
    rows.push([], headers);
    


    styles.push(
      // styles can be applied to row, col or cell, but if you do it more then once, only the latest will apply
      {col: '0-' + lastCol, width: 20},
      {row: rows.length - 1, height: 20},
      {row: 0, col: lastCol, format: 'm-d-yyyy h:mm:ss AM/PM', font: 'Arial 8 B'},
      {row: 1, col: 1, font: 'Arial 9 #000000 B', wrapText: 1},
      {row: 3 , col: 1, font: 'Arial 9 #000000 B', wrapText: 1,},
      {row: 5 , col: 1, font: 'Arial 9 #000000 B', wrapText: 1,},
      {row: rows.length - 1, col: '0-' + lastCol, font: 'Arial 10 #000000 B', wrapText: 1,},
    );
    for(let i= 0; i < 10 ;i++) {
      rows.push(['A', 'B', 'C', 'D']);
    }
    excelObj.generate(fileName, [{sheetTitle: 'Sheet 1', rows: rows, styles: styles}])

  }

  return (
    <div className="App">
      {/* <button style={{width: 200, height: 200, background: 'green'}} onClick={generateExcel}> 
      Generate Excel
      </button> */}

      <button style={{width: 200, height: 200, background: 'green'}} onClick={generateReportTemplateExcel}> 
      Generate Report Template
      </button>
    </div>
  );
}

export default App;
