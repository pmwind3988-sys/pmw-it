// The browser bundle of ExcelJS carries no typings of its own; it is the same API.
declare module 'exceljs/dist/exceljs.min.js' {
  export * from 'exceljs';
}
