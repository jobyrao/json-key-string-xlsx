'use strict';
const XLSX = require('xlsx');

class XLSX2JSON {
	/**
	 *
	 * @param source filepath or buffer
	 * @param options {object} entry: sheetName, sheets: ['sheetName']
	 * @returns {[{sheetName,data}]} sheet list
	 */
	parse(source, options = {}) {
		const readMethod = typeof source === 'string' ? 'readFile' : 'read';
		const workbook = XLSX[readMethod](source, options);
		const result = [];
		const SheetNames = workbook.SheetNames;
		if (!options.entry || !SheetNames.includes(options.entry)) {
		  options.entry = SheetNames[0];
    }
		let parsingSheetNames = options.sheets || SheetNames;
		if (parsingSheetNames instanceof Array) {
      parsingSheetNames.unshift(options.entry);
      parsingSheetNames = [...new Set(parsingSheetNames)];
    }
		for (let i of parsingSheetNames) {
			const sheetData = workbook.Sheets[i];
			result.push({
				sheetName: i,
				data: XLSX.utils.sheet_to_json(sheetData, {header: 1, raw: true, cellDates:true})
			})
		}
		for (let i = 0, len = result.length; i < len; i += 1) {
			const sheetData = result[i].data;
			for (let j = 0; j < sheetData.length; j += 1) {
				if (sheetData[j].length === 0) {
					sheetData.splice(j, 1);
					j -= 1;
				}
			}
		}
		if (/�/.test(JSON.stringify(result))) throw new Error('Exceptional characters are found!');

		return result;
	}
	parse2json(source, options = {}) {
		// 解析前，对原先数据校零处理。
		this.zeroCorrection();
		this.parsedXlsxData = this.parse(source, options);

		XLSX2JSON.switch2customStructure(this.parsedXlsxData);
		const parsedJson = this.convertProcess(this.parsedXlsxData[0].data);

		return parsedJson;
	}
	zeroCorrection() {
		// 自定义解析后的缓存
		this.parse2jsonDataCache = [];
		// 解析过程，记录一些覆盖赋值的情况。
		this.parse2jsonCover = new Set();
	}
	/**
   *
	 * Rotate each'sheet'structure. Fill in the blank'cell'
	 * @param parsedData
   * @returns undefined
	 */
	static switch2customStructure(parsedData) {
		for (let i = 0, len = parsedData.length; i < len; i += 1) {
			parsedData[i].data = XLSX2JSON.rotateMatrix(parsedData[i].data);
			parsedData[i].data[0] = XLSX2JSON.fillCellMerge(parsedData[i].data[0]);
		}
	}
	/**
	 *
	 * @param matrix
	 * @returns {[[]]}
	 */
	static rotateMatrix(matrix) {
		const results = [];
		// 首行出现require写法，其他列都是空的。旋转时需注意该列未必是空的列，可能只是require那行是空的。
		const colLen = [];
    matrix.forEach(line => {
      colLen.push(line.length);
    });
    const maxColLen = Math.max.apply(null, colLen);

		for (let i = 0, len = maxColLen; i < len; i += 1) {
			const result = [];
			for (let j = 0, lenJ = matrix.length; j < lenJ; j += 1) {
				result[j] = matrix[j][i];
			}
			results.push(result);
		}
		return results;
	}

	/**
	 * 第一例某些属性描述可能是空白，则用上一行描述值补充。如arr[],下面几行可以空白。
	 * @param firstCol
	 * @returns {*}
	 */
	static fillCellMerge(firstCol) {
		for (let i = 0, len = firstCol.length; i < len; i += 1) {
			if (firstCol[i] === undefined || (firstCol[i].trim && firstCol[i].trim() === '')) {
				if (i === 0) {
					console.warn('The first line of the json-attribute description value is forbidden to be null.');
          firstCol[i] = 'firstLineAttrDesc';
				} else {
          firstCol[i] = firstCol[i - 1];
        }
			} else {
				firstCol[i] = firstCol[i].toString().replace(/\s/g, '');
			}
		}
		return firstCol
	}

	/**
	 *
	 * @param attributeDescriptionSplited 如info[0].title 经.拆分出来的一段段。
	 * @returns {[boolean, *, null]}
	 */
	analysisAttrDesc(attributeDescriptionSplited) {
		const match = attributeDescriptionSplited.match(/\[(\d+)\]$/);
		let keyPartDescription = [];
		if (Array.isArray(match)) {
			keyPartDescription = [true, attributeDescriptionSplited.split('[')[0], + match[1]];
		} else if (attributeDescriptionSplited.slice(-2) === '[]') {
			keyPartDescription = [true, attributeDescriptionSplited.slice(0, -2), null];
		} else {
			keyPartDescription = [false, attributeDescriptionSplited, null];
		}
		return keyPartDescription;
	}
	createJsonEnumCol(columnObject, key, value, colIndex, rowIndex, sheetIndex) {
		if (key === '') return ;
		if (typeof key !== 'object') {
			key = key.split('.');
		}
		if (key.length === 0) {
			return columnObject;
		}
		let keyPart = key.shift();
		// 是否有require写法
		const requireExec = XLSX2JSON.isRequire(keyPart);
		if (requireExec) {
			// 去除#require关键字后的部分
			keyPart = requireExec[1];
			const requireSheetIndex = this.getParsedXlsxDataIndex(requireExec[2]);
			let requireSheetResult = this.parsedXlsxData[requireSheetIndex].data;

			requireSheetResult = this.convertProcess(requireSheetResult, requireSheetIndex)[colIndex];
			return this.createJsonEnumCol(columnObject, keyPart, requireSheetResult, colIndex, rowIndex, requireSheetIndex);
		}
		const keyFirstPartDesc = this.analysisAttrDesc(keyPart);
		const keyPartIsArr = keyFirstPartDesc[0];
		const keyPartKeyName = keyFirstPartDesc[1];
		const keyPartKeyIndex = keyFirstPartDesc[2];
		const isKeyPartEnd = key.length === 0;

		// arr[]、arr[0]
    if (keyPartIsArr) {
      // arr属性第一次见，先初始化
      if (columnObject[keyPartKeyName] === undefined) {
        columnObject[keyPartKeyName] = [];
      } else if (!Array.isArray(columnObject[keyPartKeyName])) {
        // arr属性已存在，但不是数组，则需要统计覆盖。并做初始化
        this.collectCoverKey(sheetIndex, rowIndex);
        columnObject[keyPartKeyName] = [];
      }
      // arr[]写法，无指定index
      if (keyPartKeyIndex === null) {
        columnObject[keyPartKeyName].push(isKeyPartEnd ? (value || '') : {});
      } else {
        // arr[0]写法，指定index
        // 数组[]在最后，其中arr[0]已经存在
        if (isKeyPartEnd && columnObject[keyPartKeyName][keyPartKeyIndex] !== undefined) {
          this.collectCoverKey(sheetIndex, rowIndex);
        }
        // 数组[]在最后
        if (isKeyPartEnd) {
          columnObject[keyPartKeyName][keyPartKeyIndex] = value || '';
        } else {
          const oldValue = columnObject[keyPartKeyName][keyPartKeyIndex];
          const source = typeof oldValue === 'object' ? oldValue : {};
          columnObject[keyPartKeyName][keyPartKeyIndex] = Object.assign({}, source);
        }
      }
      const quoteIndex = keyPartKeyIndex || columnObject[keyPartKeyName].length - 1;
      return this.createJsonEnumCol(columnObject[keyPartKeyName][quoteIndex], key, value, colIndex, rowIndex, sheetIndex);
    } else {
      // key写法
      // 当前属性key不存在
      if (columnObject[keyPartKeyName] === undefined) {
        columnObject[keyPartKeyName] = isKeyPartEnd ? (value || '') : {};
      } else if (Object.prototype.toString.call(columnObject[keyPartKeyName]) !== '[object Object]') {
        // 当前属性key存在，但不是对象，记录覆盖情况
        this.collectCoverKey(sheetIndex, rowIndex);
        columnObject[keyPartKeyName] = isKeyPartEnd ? (value || '') : {};
      } else if (isKeyPartEnd) {
        this.collectCoverKey(sheetIndex, rowIndex);
        columnObject[keyPartKeyName] = value || '';
      }
      return this.createJsonEnumCol(columnObject[keyPartKeyName], key, value, colIndex, rowIndex, sheetIndex);
    }
	}
	convertProcess(sheetData, sheetIndex = 0) {
		const parsedJson = [];
		if (this.parse2jsonDataCache[sheetIndex]) {
			return this.parse2jsonDataCache[sheetIndex];
		}
		// json属性描述列。
		const attrDescArr = sheetData[0].concat();
		for (let i = 1, len = sheetData.length; i < len; i += 1) {
			// 取出sheet第i列，语言文案
			const colValueArr = sheetData[i];
			// 该列对应的临时对象
			const colValueJson = {};
			for (let j = 0, lenJ = colValueArr.length; j < lenJ; j += 1) {

				// 返回此次tab[].c 数显转换后的深度结构的属性引用。
				this.createJsonEnumCol(colValueJson, attrDescArr[j],  colValueArr[j], i - 1, j, sheetIndex);
			}
			parsedJson.push(colValueJson);
		}
		// 记录缓存
		this.parse2jsonDataCache[sheetIndex] = parsedJson;
		return parsedJson;
	}
	static isRequire(key) {
		return /(.*)\#require\([\'\"]*([^\'\"]+)[\'\"]*\)/gi.exec(key);
	}
	getParsedXlsxDataIndex(indexOrSheetName) {
		let index;
		if (typeof +indexOrSheetName === 'number' && +indexOrSheetName < this.parsedXlsxData.length) {
			return +indexOrSheetName;
		}
		for (let i = 0, len = this.parsedXlsxData.length; i < len; i += 1) {
			if (indexOrSheetName === this.parsedXlsxData[i].sheetName) {
				index = i;
				break;
			}
		}

		if (index === undefined) {
			throw new Error(`'required' parameter is not valid: There is no such 'sheet': ${indexOrSheetName}`);
		}
		return index;
	}
  collectCoverKey(sheetIndex, rowIndex) {
	  const sheetName = this.parsedXlsxData[sheetIndex].sheetName;
	  const keyDescName = this.parsedXlsxData[sheetIndex].data[0][rowIndex];
	  const msg = `sheet name "${sheetName}", row ${rowIndex + 1}, value "${keyDescName}"`;
	  this.parse2jsonCover.add(msg);
  }
  json2XlsxByKey(jsonData, outputPath) {
	  this.keyStack = [];
	  let result = [];
	  if (Array.isArray(jsonData)) {
      result = this.object2XlsxByKey(...jsonData);
    } else {
	    result = this.object2XlsxByKey(jsonData);
    }
    if (outputPath) {
      const sheetData = XLSX.utils.aoa_to_sheet(result);
      const workbook = {
        Sheets: {
          Sheet1: sheetData
        },
        SheetNames: ['Sheet1']
      }
      if (typeof outputPath === 'string') {
        XLSX.writeFile(workbook, outputPath);
      }
      if (outputPath.type === 'array') {
        const ab = XLSX.write(workbook, { bookType: 'xlsx', bookSST: false, type: 'array' });
        return ab;
      }
    }
    return result;
  }
  object2XlsxByKey(...objs) {
    const result = [];
    const objFirst = objs[0];
    if (Array.isArray(objFirst)) {
      for (let i = 0, len = objFirst.length; i < len; i++) {
        if (typeof objFirst[i] !== 'object') {
          const preKeys = this.getPreKeyStr();
          result.push([ `${preKeys}[${i}]`, ...objs.map(obj => obj[i]) ]);
        } else {
          this.keyStack.push({
            keyName: `[${i}]`,
            type: 'array'
          })
          result.push(...this.object2XlsxByKey(...objs.map(obj => obj[i])));
        }
      }
      this.keyStack.pop();
    } else {
      const keys = Object.keys(objFirst);
      for (let i = 0, len = keys.length; i < len; i++) {
        if (typeof objFirst[keys[i]] !== 'object') {
          const preKeys = this.getPreKeyStr();
          result.push([ (preKeys && preKeys + '.') + keys[i], ...objs.map(obj => obj[keys[i]]) ]);
        } else {
          this.keyStack.push({
            keyName: keys[i],
            type: 'object'
          })
          result.push(...this.object2XlsxByKey(...objs.map(obj => obj[keys[i]])));
        }
      }
      this.keyStack.pop();
    }
    return result;
  }
  getPreKeyStr() {
    let preKeyStr = '';
    if (!this.keyStack.length) {
      return preKeyStr;
    }
    for (let i = 0, len = this.keyStack.length; i < len; i++) {
      if (i < len - 1) {
        if (this.keyStack[i + 1].type === 'object') {
          preKeyStr += this.keyStack[i].keyName + '.';
        } else {
          preKeyStr += this.keyStack[i].keyName;
        }
      } else {
        preKeyStr += this.keyStack[i].keyName;
      }
    }
    return preKeyStr;
  }
}
// Compatible with syntax of versions before v0.1.0
const defaultCase = new XLSX2JSON();
Object.defineProperties(XLSX2JSON, {
  parse: {
    value: defaultCase.parse.bind(defaultCase)
  },
  parse2json: {
    value: defaultCase.parse2json.bind(defaultCase)
  },
  parse2jsonDaTa: {
    value: defaultCase.parse2jsonDaTa
  },
  parse2jsonCover: {
    value: defaultCase.parse2jsonCover
  },
  parsedXlsxData: {
    value: defaultCase.parsedXlsxData
  }
});

module.exports = XLSX2JSON;
