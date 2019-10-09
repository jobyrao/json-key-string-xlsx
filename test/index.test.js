'use strict';

const chai = require('chai');
const xlsx2json = require('../lib/index');
const path = require('path');

const expect = chai.expect;
const testXlsxPath = path.join(__dirname, './assets/excel.xlsx');

describe('原生数据结构parse解析测试', function() {
  const parsedNativeData = xlsx2json.parse(testXlsxPath);
  const parsedNativeData2 = xlsx2json.parse(testXlsxPath, {entry: 'company', sheets: ['address']});
  it('原生解析：sheet个数完整', function() {
    expect(parsedNativeData.length).to.be.equal(4);
  });
  it('原生解析：sheet数据结构完整', function() {
    expect(parsedNativeData[0].sheetName).to.be.equal('main');
    expect(parsedNativeData[0].data[0][0]).to.be.equal('filename');
    expect(parsedNativeData[3].data[0][0]).to.be.equal('city[]');
  });
  it('原生解析：指定解析入口', function() {
    expect(parsedNativeData2[0].sheetName).to.be.equal('company');
  });
  it('原生解析：指定解析sheets', function() {
    expect(parsedNativeData2.length).to.be.equal(2);
    expect(parsedNativeData2[0].sheetName).to.be.equal('company');
    expect(parsedNativeData2[1].sheetName).to.be.equal('address');
  });
});

describe('自定义数据结构parse2json解析测试', function() {
  const parsedCustomData = xlsx2json.parse2json(testXlsxPath);
  const parsedCustomData2 = xlsx2json.parse2json(testXlsxPath, {entry: 'company', sheets: ['address']});
  it('自定义解析：语言码个数完整', function() {
    expect(parsedCustomData.length).to.be.equal(2);
  });
  it('自定义解析：第一层字段', function() {
    expect(parsedCustomData[0].filename).to.be.equal('cn');
    expect(parsedCustomData[1].filename).to.be.equal('en');
  });
  it('自定义解析：数组结构', function() {
    expect(parsedCustomData[0].disclaimer.content[0]).to.be.equal('1，非人工检索方式');
    expect(parsedCustomData[1].disclaimer.content[0]).to.be.equal('1. Non-manual retrieval');
  });
  it('自定义解析：深度数据结构', function() {
    expect(parsedCustomData[0].disclaimer.content[1].c.a[0].b[0]).to.be.equal('6，网络传播权');
    expect(parsedCustomData[1].disclaimer.content[1].c.a[0].b[0]).to.be.equal('6. Right of Network Communication');
  });
  it('自定义解析：主入口第一个sheet中#require结构', function() {
    expect(parsedCustomData[0].statusCode['200']).to.be.equal('成功');
    expect(parsedCustomData[1].statusCode['200']).to.be.equal('Success');
  });
  it('自定义解析：其他sheet中#require结构', function() {
    expect(parsedCustomData[0].companyInfo.address.city[0]).to.be.equal('广州市');
    expect(parsedCustomData[1].companyInfo.address.city[0]).to.be.equal('guangzhou');
  });
  it('自定义解析：指定入口和sheets', function() {
    expect(parsedCustomData2[0].address.city[0]).to.be.equal('广州市');
    expect(parsedCustomData2[1].address.city[0]).to.be.equal('guangzhou');
  });
  it('自定义解析：a[0].name,a[0].nick数组项对象补充写法', function() {
    expect(parsedCustomData2[0].group[0].name).to.be.equal('蚂蚁');
    expect(parsedCustomData2[0].group[0].address).to.be.equal('深圳');
    expect(parsedCustomData2[0].group.length).to.be.equal(2);
  });
});
