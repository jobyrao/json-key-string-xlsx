import {ParsingOptions} from 'xlsx';

interface Options extends ParsingOptions{
  entry?: string;
  sheets?: string[];
}
interface NativeData {
  sheetName: string;
  data: string[][];
}
interface Lang {
  [key: string]: any;
}
interface WriteOptions {
  type: string;
}

export default class XLSX2JSON {
  parse(data: any, options?: Options): NativeData[];
  parse2json(data: any, options?: Options): Lang[];
  parse2jsonCover: string[];
  json2XlsxByKey(data: Lang | Lang[], options?: string | WriteOptions): string[][] | ArrayBuffer;
}
