# xlsx-json-js

[![build status](http://img.shields.io/travis/diyao/xlsx-json-js/master.svg?style=flat)](http://travis-ci.org/diyao/xlsx-json-js)
[![Coverage Status](https://coveralls.io/repos/diyao/xlsx-json-js/badge.svg?branch=)](https://coveralls.io/r/diyao/xlsx-json-js?branch=master)
[![npm version](https://img.shields.io/npm/v/xlsx-json-js.svg?style=flat)](https://www.npmjs.com/package/xlsx-json-js)
[![license](https://img.shields.io/github/license/diyao/xlsx-json-js.svg)](https://tldrlegal.com/license/mit-license)

## Install
```bash
$ npm i xlsx-json-js --save
```
## Usage
The contents of the sample file `excel.xlsx` are as follows.
[google docs](https://docs.google.com/spreadsheets/d/18BDeB2zNKA2AuMFDMcJuIdHBDaNRskYQPmZIv_1A5p0/edit#gid=1308189912) 
or 
[qq docs](https://docs.qq.com/sheet/DY0JTcGNjT3NFcWNw)
### 1. Parsing into native structure

```javascript
const xlsx2json = require('xlsx-json-js');
const path = require('path');
const xlsxPath = path.join('./excel.xlsx');
// filepath or buffer
const nativeData = xlsx2json.parse(xlsxPath);
```

<details>
  <summary><b>console.log(nativeData)</b> (click to show)</summary>

```json
[
  {
    "sheetName": "main",
    "data": [
      [
        "filename",
        "cn",
        "en"
      ],
      [
        "lang",
        "cn",
        "en"
      ],
      [
        "title",
        "文章标题",
        "Title of article"
      ],
      [
        "userInfo[0].name",
        "用户名",
        "username"
      ],
      [
        "userInfo[1].nickname",
        "用户昵称",
        "nickname"
      ],
      [
        "disclaimer.title",
        "免责申明",
        "Disclaimer"
      ],
      [
        "disclaimer.content[]",
        "1，非人工检索方式",
        "1. Non-manual retrieval"
      ],
      [
        null,
        "2，搜索链接到的第三方网页",
        "2. Search for linked third-party pages"
      ],
      [
        null,
        "3，自动搜索获得",
        "3. Automatic Search Acquisition"
      ],
      [
        " ",
        "4，自行承担风险",
        "4. Take risks on your own"
      ],
      [
        "disclaimer.content[]",
        "5，个人隐私权",
        "5. Right to personal privacy"
      ],
      [
        "disclaimer.content[1].c.a[0].b[0]",
        "6，网络传播权",
        "6. Right of Network Communication"
      ],
      [
        "statusCode#require(2)"
      ],
      [
        "companyInfo#require(company)"
      ],
      [
        "for test1.key",
        "test1 key值，原始",
        "test1 key value, native"
      ],
      [
        "for test1.key",
        "test1 key值，再次",
        "test1 key value, again"
      ],
      [
        "for test1.key2",
        "test1 key2值",
        "test1 key2 value"
      ],
      [
        "fortest2[].key",
        "test2 key值，原始",
        "test2 key value, native"
      ],
      [
        "fortest2[].key",
        "test2 key值，再次",
        "test2 key value, again"
      ],
      [
        "for test2[].key2",
        "test2 key2值",
        "test2 key2 value"
      ],
      [
        "for test2[1].key2",
        "test2 key值，覆盖",
        "test2 key value, cover"
      ],
      [
        "fortest3.key",
        "test3 对象",
        "test3 object"
      ],
      [
        "fortest3[].key",
        "test3 对象改为数组",
        "test3 object to array"
      ],
      [
        "fortest4[].key",
        "test4 数组",
        "test4 array"
      ],
      [
        "fortest4.key",
        "test4 数组改为对象",
        "test4 array to object"
      ]
    ]
  },
  {
    "sheetName": "company",
    "data": [
      [
        "name",
        "公司名",
        "company name"
      ],
      [
        "address#require(address)"
      ],
      [
        "industry",
        "互联网",
        "Internet"
      ]
    ]
  },
  {
    "sheetName": "statusCode",
    "data": [
      [
        200,
        "成功",
        "Success"
      ],
      [
        404,
        "失败",
        "fail"
      ]
    ]
  },
  {
    "sheetName": "address",
    "data": [
      [
        "city[]",
        "广州市",
        "guangzhou"
      ],
      [
        "address[]",
        "番禺万达",
        "Wanda, Panyu District, Guangzhou"
      ],
      [
        "city[]",
        "北京市",
        "beijing"
      ],
      [
        "address[]",
        "某大厦",
        "Zhizhen Building"
      ]
    ]
  }
]
```

</details>

### 2. Resolve to a custom structure

```JavaScript
const xlsx2json = require('xlsx-json-js');
const path = require('path');
const xlsxPath = path.join('./excel.xlsx');

const customData = xlsx2json.parse2json(xlsxPath);
// console.log(xlsx2json.parse2jsonDataCache);
// console.log(xlsx2json.parse2jsonCover);
// console.log(xlsx2json.parsedXlsxData);
```

<details>
  <summary><b>console.log(customData)</b> (click to show)</summary>

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

<details>
  <summary><b>console.log([...xlsx2json.parse2jsonCover])</b> (click to show)</summary>

```text
[ 'sheet name "main", row 12, value "disclaimer.content[1].c.a[0].b[0]"',
  'sheet name "main", row 16, value "fortest1.key"',
  'sheet name "main", row 21, value "fortest2[1].key2"',
  'sheet name "main", row 23, value "fortest3[].key"',
  'sheet name "main", row 25, value "fortest4.key"' ]
```

</details>

## License

[MIT](LICENSE)
