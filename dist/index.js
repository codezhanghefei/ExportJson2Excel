(function (global, factory) {
  typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory(require('file-saver'), require('xlsx')) :
  typeof define === 'function' && define.amd ? define(['file-saver', 'xlsx'], factory) :
  (global = typeof globalThis !== 'undefined' ? globalThis : global || self, global.index = factory(global.saveAs, global.XLSX));
})(this, (function (saveAs, XLSX) { 'use strict';

  function _interopDefaultLegacy (e) { return e && typeof e === 'object' && 'default' in e ? e : { 'default': e }; }

  var saveAs__default = /*#__PURE__*/_interopDefaultLegacy(saveAs);
  var XLSX__default = /*#__PURE__*/_interopDefaultLegacy(XLSX);

  /* eslint-disable */

  function datenum(v, date1904) {
    if (date1904) v += 1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
  }

  function sheet_from_array_of_arrays(data, opts) {
    var ws = {};
    var range = {
      s: {
        c: 10000000,
        r: 10000000
      },
      e: {
        c: 0,
        r: 0
      }
    };
    for (var R = 0; R != data.length; ++R) {
      for (var C = 0; C != data[R].length; ++C) {
        if (range.s.r > R) range.s.r = R;
        if (range.s.c > C) range.s.c = C;
        if (range.e.r < R) range.e.r = R;
        if (range.e.c < C) range.e.c = C;
        var cell = {
          v: data[R][C]
        };
        if (cell.v == null) continue;
        var cell_ref = XLSX__default["default"].utils.encode_cell({
          c: C,
          r: R
        });

        if (typeof cell.v === 'number') cell.t = 'n';
        else if (typeof cell.v === 'boolean') cell.t = 'b';
        else if (cell.v instanceof Date) {
          cell.t = 'n';
          cell.z = XLSX__default["default"].SSF._table[14];
          cell.v = datenum(cell.v);
        } else cell.t = 's';

        ws[cell_ref] = cell;
      }
    }
    if (range.s.c < 10000000) ws['!ref'] = XLSX__default["default"].utils.encode_range(range);
    return ws;
  }

  function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
  }

  function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }

  function formatJson(filterVal, jsonData) {
    return jsonData.map(v => filterVal.map(j => {
      return v[j];
    }));
  }

  function export_json_to_excel({
    multiHeader = [],
    header,
    data,
    filename,
    noteList = [],
    merges = [],
    autoWidth = true,
    bookType = 'xlsx'
  } = {}) {
    /* original data */
    filename = filename || "excel-list";
    data = [...data];
    data.unshift(header);

    for (let i = multiHeader.length - 1; i > -1; i--) {
      data.unshift(multiHeader[i]);
    }

    let dataStart = 0;
    if (noteList.length) {
      data.unshift([""]);
      data.unshift(noteList);
      dataStart = noteList.length + 1;
    }

    var ws_name = "SheetJS";
    var wb = new Workbook(),
      ws = sheet_from_array_of_arrays(data);

    if (merges.length > 0) {
      if (!ws["!merges"]) ws["!merges"] = [];
      merges.forEach((item) => {
        ws["!merges"].push(XLSX__default["default"].utils.decode_range(item));
      });
    }

    if (autoWidth) {
      /*设置worksheet每列的最大宽度*/
      const colWidth = data.slice(dataStart).map((row) =>
        row.map((val) => {
          /*先判断是否为null/undefined*/
          if (val == null) {
            return {
              wch: 10,
            };
          } else if (val.toString().charCodeAt(0) > 255) {
          /*再判断是否为中文*/
            return {
              wch: val.toString().length * 2,
            };
          } else {
            return {
              wch: val.toString().length,
            };
          }
        })
      );
      /*以第一行为初始值*/
      let result = colWidth[0];
      for (let i = 1; i < colWidth.length; i++) {
        for (let j = 0; j < colWidth[i].length; j++) {
          if (result[j]["wch"] < colWidth[i][j]["wch"]) {
            result[j]["wch"] = colWidth[i][j]["wch"];
          }
        }
      }
      ws["!cols"] = result;
    }

    /* add worksheet to workbook */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;

    var wbout = XLSX__default["default"].write(wb, {
      bookType: bookType,
      bookSST: false,
      type: "binary",
    });
    saveAs__default["default"](
      new Blob([s2ab(wbout)], {
        type: "application/octet-stream",
      }),
      `${filename}.${bookType}`
    );
  }

  /**
   * 将数组导出为excel
   * @param {{title:string,dataIndex:string}[]} fields 要导出的字段 [{ title, dataIndex }]
   * @param {{dataIndex:any}[]} data 要导出的数据，是个数组
   * @param {string} filename 导出的文件名称
   * @param {?string[]} noteList 备注，如有则放在表格开头行 
   */
  function exportJson2Excel(fields, data, filename,noteList=[]) {
    const exportData = formatJson(fields.map((item) => item.dataIndex), data);

    export_json_to_excel({
      header: fields.map((item) => item.title),
      data: exportData,
      filename,
      noteList:noteList
    });
  }

  return exportJson2Excel;

}));
