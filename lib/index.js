'use strict';
const XLSX = require('xlsx');

class XLSX2JSON {
	constructor() {
	}
	/**
	 *
	 * @param source filepath or buffer
	 * @param options
	 * @returns {[{sheetName,data}]} sheet list
	 */
	parse(source, options = {}) {
		const readMethod = typeof source === 'string' ? 'readFile' : 'read';
		const workbook = XLSX[readMethod](source, options);
		const result = [];
		for (let i in workbook.Sheets) {
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
		if (/�/.test(JSON.stringify(result))) {
			console.error('Exceptional characters are found after parsed!');
			process.exit(0);
		}

		return result;
	}
	parse2json(source, options = {}) {
		// 解析前，对原先数据校零处理。
		this.zeroCorrection();
		const opts = Object.assign({isColOriented: true}, options);
		this.parsedXlsxData = this.parse(source, options);

		XLSX2JSON.switch2customStructure(this.parsedXlsxData);
		const parsedJson = this.convertProcess(this.parsedXlsxData[0].data);

		return parsedJson;
	}
	zeroCorrection() {
		// 自定义解析后的缓存
		this.parse2jsonDataCache = [];
		// 解析过程，记录一些覆盖赋值的情况。
		this.parse2jsonCover = [];
	}
	/**
	 * Rotate each'sheet'structure. Fill in the blank'cell'
	 * @param parsedData
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
		for (let i = 0, len = matrix[0].length; i < len; i += 1) {
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
					console.error('The first line of the\'json\'attribute description value is forbidden to be null.');
					process.exit(0);
				}
				firstCol[i] = firstCol[i - 1];
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
		}else if (attributeDescriptionSplited.slice(-2) === '[]') {
			keyPartDescription = [true, attributeDescriptionSplited.slice(0, -2), null];
		}else{
			keyPartDescription = [false, attributeDescriptionSplited, null];
		}
		return keyPartDescription;
	}
	createJsonEnumCol(columnObject, key, value, colIndex) {
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
			return this.createJsonEnumCol(columnObject, keyPart, requireSheetResult, colIndex);
		}
		const keyFirstPartDesc = this.analysisAttrDesc(keyPart);
		const keyPartIsArr = keyFirstPartDesc[0];
		const keyPartKeyName = keyFirstPartDesc[1];
		const keyPartKeyIndex = keyFirstPartDesc[2];
		const isKeyPartEnd = key.length === 0;

    if (keyPartIsArr) {
      if (columnObject[keyPartKeyName] === undefined) {
        columnObject[keyPartKeyName] = [];
      } else if (!Array.isArray(columnObject[keyPartKeyName])) {
        this.parse2jsonCover.push(`${XLSX2JSON.getJsonCoverKey(keyPart, key)}`);
        columnObject[keyPartKeyName] = [];
      }
      if (keyPartKeyIndex === null) {
        columnObject[keyPartKeyName].push(isKeyPartEnd ? value : {});
      } else {
        if (columnObject[keyPartKeyName][keyPartKeyIndex] !== undefined) {
          this.parse2jsonCover.push(`${XLSX2JSON.getJsonCoverKey(keyPart, key)}`);
        }
        columnObject[keyPartKeyName][keyPartKeyIndex] = isKeyPartEnd ? value : {};
      }
      const quoteIndex = keyPartKeyIndex || columnObject[keyPartKeyName].length - 1;
      return this.createJsonEnumCol(columnObject[keyPartKeyName][quoteIndex], key, value, colIndex);
    } else {
      if (columnObject[keyPartKeyName] === undefined) {
        columnObject[keyPartKeyName] = isKeyPartEnd ? value : {};
      } else if (Object.prototype.toString.call(columnObject[keyPartKeyName]) !== '[object Object]') {
        this.parse2jsonCover.push(`${XLSX2JSON.getJsonCoverKey(keyPart, key)}`);
        columnObject[keyPartKeyName] = isKeyPartEnd ? value : {};
      }else {
        if (isKeyPartEnd) {
          this.parse2jsonCover.push(`${XLSX2JSON.getJsonCoverKey(keyPart, key)}`);
          columnObject[keyPartKeyName] = value;
        }
      }
      return this.createJsonEnumCol(columnObject[keyPartKeyName], key, value, colIndex);
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
				this.createJsonEnumCol(colValueJson, attrDescArr[j],  colValueArr[j], i - 1);
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
			console.error(`'required' parameter is not valid: There is no such 'sheet': ${indexOrSheetName}`);
			process.exit(0);
		}
		return index;
	}
	static getJsonCoverKey(keyPart, key) {
	  return `${keyPart}${key.length ? '.' + key.join('.') : ''}`;
  }
}

module.exports = new XLSX2JSON();
