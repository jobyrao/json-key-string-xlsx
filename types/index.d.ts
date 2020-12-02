import {ParsingOptions} from 'xlsx';

interface Options extends ParsingOptions{
  entry?: string;
}
interface NativeData {
  sheetName: string;
  data: string[][];
}
interface Lang {
  [key: string]: any;
}

declare module 'xlsx-json-js' {
  class XLSX2JSON {
    parse(data: any, options?: Options): NativeData[];
    parse2json(data: any, options?: Options): Lang[];
    parse2jsonCover: string[];
  }

  export default XLSX2JSON
}
