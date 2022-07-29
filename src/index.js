/* global window, document */
import { h } from './component/element';
import DataProxy from './core/data_proxy';
import helper from './core/helper';
import Sheet from './component/sheet';
import Bottombar from './component/bottombar';
import { cssPrefix } from './config';
import { locale } from './locale/locale';
import './index.less';
import { CellRange } from './core/cell_range';


class Spreadsheet {
  constructor(selectors, options = {}) {
    let targetEl = selectors;
    this.options = { showBottomBar: true, ...options };
    this.sheetIndex = 1;
    this.activeSheetIndex = 0;
    this.datas = [];
    if (typeof selectors === 'string') {
      targetEl = document.querySelector(selectors);
    }
    this.bottombar = this.options.showBottomBar ? new Bottombar(() => {
      if (this.options.mode === 'read') return;
      const d = this.addSheet();
      this.sheet.resetData(d);
    }, (index) => {
      const d = this.datas[index];
      this.sheet.resetData(d);
      this.activeSheetIndex = index;
    }, () => {
      this.deleteSheet();
    }, (index, value) => {
      this.datas[index].name = value;
      this.sheet.trigger('change');
    }) : null;
    this.data = this.addSheet();
    const rootEl = h('div', `${cssPrefix}`)
      .on('contextmenu', evt => evt.preventDefault());
    // create canvas element
    targetEl.appendChild(rootEl.el);
    this.sheet = new Sheet(rootEl, this.data);
    if (this.bottombar !== null) {
      rootEl.child(this.bottombar.el);
    }
  }

  addSheet(name, active = true) {
    const n = name || `sheet${this.sheetIndex}`;
    const d = new DataProxy(n, this.options);
    d.change = (...args) => {
      this.sheet.trigger('change', ...args);
    };
    this.datas.push(d);
    // console.log('d:', n, d, this.datas);
    if (this.bottombar !== null) {
      this.bottombar.addItem(n, active, this.options);
    }
    this.sheetIndex += 1;
    return d;
  }

  deleteSheet() {
    if (this.bottombar === null) return;

    const [oldIndex, nindex] = this.bottombar.deleteItem();
    if (oldIndex >= 0) {
      this.datas.splice(oldIndex, 1);
      if (nindex >= 0) this.sheet.resetData(this.datas[nindex]);
      this.sheet.trigger('change');
    }
  }

  loadData(data) {
    const ds = Array.isArray(data) ? data : [data];
    if (this.bottombar !== null) {
      this.bottombar.clear();
    }
    this.datas = [];
    if (ds.length > 0) {
      for (let i = 0; i < ds.length; i += 1) {
        const it = ds[i];
        const nd = this.addSheet(it.name, i === 0);
        nd.setData(it);
        if (i === 0) {
          this.sheet.resetData(nd);
        }
      }
    }
    return this;
  }

  getData() {
    return this.datas.map(it => it.getData());
  }

  cellText(ri, ci, text, sheetIndex = 0) {
    this.datas[sheetIndex].setCellText(ri, ci, text, 'finished');
    return this;
  }

  cell(ri, ci, sheetIndex = 0) {
    return this.datas[sheetIndex].getCell(ri, ci);
  }

  cellStyle(ri, ci, sheetIndex = 0) {
    return this.datas[sheetIndex].getCellStyle(ri, ci);
  }

  setCellStyle(ri, ci, property, value, sheetIndex = 0) {
    const dataProxy = this.datas[sheetIndex];
    const { styles, rows } = dataProxy;
    const cell = rows.getCellOrNew(ri, ci);
    let cstyle = {};
    if (cell.style !== undefined) {
      cstyle = helper.cloneDeep(styles[cell.style]);
    }
    if (property === 'format') {
      cstyle.format = value;
      cell.style = dataProxy.addStyle(cstyle);
    } else if (
      property === 'font-bold'
      || property === 'font-italic'
      || property === 'font-name'
      || property === 'font-size'
    ) {
      const nfont = {};
      nfont[property.split('-')[1]] = value;
      cstyle.font = Object.assign(cstyle.font || {}, nfont);
      cell.style = dataProxy.addStyle(cstyle);
    } else if (
      property === 'strike'
      || property === 'textwrap'
      || property === 'underline'
      || property === 'align'
      || property === 'valign'
      || property === 'color'
      || property === 'bgcolor'
    ) {
      cstyle[property] = value;
      cell.style = dataProxy.addStyle(cstyle);
    } else {
      cell[property] = value;
    }
    return this;
  }

  setCellBorderStyle(ri, ci, mode, style, color, sheetIndex = 0) {
    const dataProxy = this.datas[sheetIndex];
    const { styles, rows } = dataProxy;
    const cell = rows.getCellOrNew(ri, ci);
    let cstyle = {};

    if (cell.style !== undefined) {
      cstyle = helper.cloneDeep(styles[cell.style]);
    }

    const bss = {};
    bss[mode] = [style, color];
    cstyle = helper.merge(cstyle, { border: bss });
    cell.style = dataProxy.addStyle(cstyle);

    return this;
  }

  mergeCellsRange(range, sheetIndex = 0) {
    return this.mergeCellsInner(CellRange.valueOf(range), sheetIndex);
  }

  mergeCells(sri, sci, eri, eci, sheetIndex = 0) {
    return this.mergeCellsInner(new CellRange(sri, sci, eri, eci), sheetIndex);
  }

  mergeCellsInner(cr, sheetIndex = 0) {
    const dataProxy = this.datas[sheetIndex];
    const cell = dataProxy.rows.getCellOrNew(cr.sri, cr.sci);
    cell.merge = [cr.eri - cr.sri, cr.eci - cr.sci];
    dataProxy.merges.add(cr);
    dataProxy.rows.deleteCells(cr);
    dataProxy.rows.setCell(cr.sri, cr.sci, cell);
    return this;
  }

  reRender() {
    this.sheet.table.render();
    return this;
  }

  on(eventName, func) {
    this.sheet.on(eventName, func);
    return this;
  }

  validate() {
    const { validations } = this.data;
    return validations.errors.size <= 0;
  }

  change(cb) {
    this.sheet.on('change', cb);
    return this;
  }

  static locale(lang, message) {
    locale(lang, message);
  }
}

const spreadsheet = (el, options = {}) => new Spreadsheet(el, options);

if (window) {
  window.x_spreadsheet = spreadsheet;
  window.x_spreadsheet.locale = (lang, message) => locale(lang, message);
}

export default Spreadsheet;
export {
  spreadsheet,
};
