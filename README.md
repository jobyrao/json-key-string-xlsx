# json-key-string-xlsx

[![npm version](https://img.shields.io/npm/v/json-key-string-xlsx.svg?style=flat)](https://www.npmjs.com/package/json-key-string-xlsx)
![GitHub Workflow Status](https://img.shields.io/github/workflow/status/jobyrao/json-key-string-xlsx/Continuous%20integration)
[![codecov](https://codecov.io/gh/jobyrao/json-key-string-xlsx/branch/master/graph/badge.svg?token=OK5M7HAAU7)](https://codecov.io/gh/jobyrao/json-key-string-xlsx)
![npms.io (quality)](https://img.shields.io/npms-io/quality-score/json-key-string-xlsx)
[![GitHub issues](https://img.shields.io/github/issues/jobyrao/json-key-string-xlsx)](https://github.com/jobyrao/json-key-string-xlsx/issues)
[![license](https://img.shields.io/github/license/jobyrao/json-key-string-xlsx.svg)](https://tldrlegal.com/license/mit-license)

## Introduction
Convert between json and xlsx files by key string in a browser or NodeJS.

## Quick Preview
### Content structure of xlsx file
1. The first column is a string of JSON key names.
2. The other columns are the values of the JSON key string.

|  |  |  |
| ---------- | -----------| -----------|
| lang              | cn   | en   |
| userInfo[0].name   | 用户名   | username   |
| userInfo[0].nickname | 昵称   | nickname   |

### Excel file will be converted to
Similarly, the JSON can also be converted to xlsx file.
```json
[
  {
    "lang": "cn",
    "userInfo": [
      {
        "name": "用户名",
        "nickname": "昵称"
      }
    ]
  },
  {
    "lang": "en",
    "userInfo": [
      {
        "name": "username",
        "nickname": "nickname"
      }
    ]
  }
]
```
## Getting Started
### Install
In the browser, just add a script tag:
```html
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.15.0/dist/xlsx.full.min.js"></script>
<script src="dist/json-key-string-xlsx.umd.min.js"></script>
```
<details>
  <summary><b>CDN Availability</b> (click to show)</summary>

|  |  |
| ---------- | -----------|
| unpkg   | https://unpkg.com/json-key-string-xlsx/ |
| jsDelivr | https://jsdelivr.com/package/npm/json-key-string-xlsx |

</details>

With npm:

```bash
$ npm i json-key-string-xlsx --save
```

### Usage

#### 1. Convert xlsx file to JSON
Sample files of this api:
[google docs](https://docs.google.com/spreadsheets/d/18BDeB2zNKA2AuMFDMcJuIdHBDaNRskYQPmZIv_1A5p0/edit#gid=1308189912)
or
[tencent docs](https://docs.qq.com/sheet/DY0JTcGNjT3NFcWNw)

Commonjs
```JavaScript
const XLSX2JSON = require('json-key-string-xlsx');
const xlsx2json = new XLSX2JSON();
const path = require('path');
const xlsxPath = path.join('./excel.xlsx');

const jsonData = xlsx2json.parse2json(xlsxPath);
// console.log(xlsx2json.parse2jsonCover);
```
ES Module
```js
import xlsxJsonJs from 'json-key-string-xlsx';
const xlsx2json = new xlsxJsonJs();
```
UMD
```html
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.15.0/dist/xlsx.full.min.js"></script>
<script src="dist/json-key-string-xlsx.umd.min.js"></script>

<input type="file" name="file" id="file">
<script type="text/javascript">
  const xlsx2json = new jsonKeyStringXlsx();
  function handleFile(e) {
    const files = e.target.files, f = files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
      const data = new Uint8Array(e.target.result);
      const jsonData = xlsx2json.parse2json(data, {type: 'array'});
    };
    reader.readAsArrayBuffer(f);
  }
  document.getElementById('file').addEventListener('change', handleFile, false);
</script>
```


<details>
  <summary><b>console.log(jsonData)</b> (click to show)</summary>

```text
[
  {
    "filename": "cn",
    "lang": "cn",
    "title": "文章标题",
    "userInfo": [
      {
        "name": "用户名"
      },
      {
        "nickname": "用户昵称"
      }
    ],
    "disclaimer": {
      "title": "免责申明",
      "content": [
        "1，非人工检索方式",
        {
          "c": {
            "a": [
              {
                "b": [
                  "6，网络传播权"
                ]
              }
            ]
          }
        },
        "3，自动搜索获得",
        "4，自行承担风险",
        "5，个人隐私权"
      ]
    },
    "statusCode": {
      "200": "成功",
      "404": "失败"
    },
    "companyInfo": {
      "name": "公司名",
      "address": {
        "city": [
          "广州市",
          "北京市"
        ],
        "address": [
          "番禺万达",
          "某大厦"
        ]
      },
      "industry": "互联网"
    },
    "fortest1": {
      "key": "test1 key值，再次",
      "key2": "test1 key2值"
    },
    "fortest2": [
      {
        "key": "test2 key值，原始"
      },
      {
        "key2": "test2 key值，覆盖"
      },
      {
        "key2": "test2 key2值"
      }
    ],
    "fortest3": [
      {
        "key": "test3 对象改为数组"
      }
    ],
    "fortest4": {
      "key": "test4 数组改为对象"
    }
  },
  {
    "filename": "en",
    ......
  }
]
```

</details>

If some keys are overwritten, you can get details from `xlsx2json.parse2jsonCover`.

```text
[ 'sheet name "main", row 12, value "disclaimer.content[1].c.a[0].b[0]"',
  'sheet name "main", row 16, value "fortest1.key"',
  'sheet name "main", row 21, value "fortest2[1].key2"',
  'sheet name "main", row 23, value "fortest3[].key"',
  'sheet name "main", row 25, value "fortest4.key"' ]
```

#### 2. Convert JSON to xlsx file
The `objData` can also be an array of object.
```js
const XLSX2JSON = require('json-key-string-xlsx');
const xlsx2json = new XLSX2JSON();
const objData = {
  "lang": "en",
  "userInfo": [
    {
      "name": "username",
      "nickname": "nickname"
    }
  ]
}
const aoaData = xlsx2json.json2XlsxByKey(objData);
// const aoaData = xlsx2json.json2XlsxByKey(objData, outputPath);
```
Output file:

|  |    |
| ---------- |  -----------|
| lang              | en   |
| userInfo[0].name   | username   |
| userInfo[0].nickname | nickname   |

## License

[MIT](LICENSE)
