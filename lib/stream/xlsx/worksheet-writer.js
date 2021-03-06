const {once} = require('events');
const _ = require('../../utils/under-dash');
const RelType = require('../../xlsx/rel-type');

const colCache = require('../../utils/col-cache');
const Dimensions = require('../../doc/range');
const StringBuf = require('../../utils/string-buf');

const Row = require('../../doc/row');
const Column = require('../../doc/column');

const SheetRelsWriter = require('./sheet-rels-writer');
const SheetCommentsWriter = require('./sheet-comments-writer');
const DataValidations = require('../../doc/data-validations');

const xmlBuffer = new StringBuf();

// ============================================================================================
// Xforms
const ListXform = require('../../xlsx/xform/list-xform');
const DataValidationsXform = require('../../xlsx/xform/sheet/data-validations-xform');
const SheetPropertiesXform = require('../../xlsx/xform/sheet/sheet-properties-xform');
const SheetFormatPropertiesXform = require('../../xlsx/xform/sheet/sheet-format-properties-xform');
const ColXform = require('../../xlsx/xform/sheet/col-xform');
const RowXform = require('../../xlsx/xform/sheet/row-xform');
const HyperlinkXform = require('../../xlsx/xform/sheet/hyperlink-xform');
const SheetViewXform = require('../../xlsx/xform/sheet/sheet-view-xform');
const SheetProtectionXform = require('../../xlsx/xform/sheet/sheet-protection-xform');
const PageMarginsXform = require('../../xlsx/xform/sheet/page-margins-xform');
const PageSetupXform = require('../../xlsx/xform/sheet/page-setup-xform');
const AutoFilterXform = require('../../xlsx/xform/sheet/auto-filter-xform');
const PictureXform = require('../../xlsx/xform/sheet/picture-xform');
const ConditionalFormattingsXform = require('../../xlsx/xform/sheet/cf/conditional-formattings-xform');

// since prepare and render are functional, we can use singletons
const xform = {
  dataValidations: new DataValidationsXform(),
  sheetProperties: new SheetPropertiesXform(),
  sheetFormatProperties: new SheetFormatPropertiesXform(),
  columns: new ListXform({tag: 'cols', length: false, childXform: new ColXform()}),
  row: new RowXform(),
  hyperlinks: new ListXform({tag: 'hyperlinks', length: false, childXform: new HyperlinkXform()}),
  sheetViews: new ListXform({tag: 'sheetViews', length: false, childXform: new SheetViewXform()}),
  sheetProtection: new SheetProtectionXform(),
  pageMargins: new PageMarginsXform(),
  pageSeteup: new PageSetupXform(),
  autoFilter: new AutoFilterXform(),
  picture: new PictureXform(),
  conditionalFormattings: new ConditionalFormattingsXform(),
};

// ============================================================================================

class WorksheetWriter {
  constructor(options) {
    // in a workbook, each sheet will have a number
    this.id = options.id;

    // and a name
    this.name = options.name || `Sheet${this.id}`;

    // add a state
    this.state = options.state || 'visible';

    // rows are stored here while they need to be worked on.
    // when they are committed, they will be deleted.
    this._rows = [];

    // column definitions
    this._columns = null;

    // column keys (addRow convenience): key ==> this._columns index
    this._keys = {};

    // keep record of all merges
    this._merges = [];
    this._merges.add = function() {}; // ignore cell instruction

    // keep record of all hyperlinks
    this._sheetRelsWriter = new SheetRelsWriter(options);

    this._sheetCommentsWriter = new SheetCommentsWriter(this, this._sheetRelsWriter, options);

    // keep a record of dimensions
    this._dimensions = new Dimensions();

    // first uncommitted row
    this._rowZero = 1;

    // committed flag
    this.committed = false;

    // for data validations
    this.dataValidations = new DataValidations();

    // for sharing formulae
    this._formulae = {};
    this._siFormulae = 0;

    // keep a record of conditionalFormattings
    this.conditionalFormatting = [];

    // for default row height, outline levels, etc
    this.properties = Object.assign(
      {},
      {
        defaultRowHeight: 15,
        dyDescent: 55,
        outlineLevelCol: 0,
        outlineLevelRow: 0,
      },
      options.properties
    );

    // for all things printing
    this.pageSetup = Object.assign(
      {},
      {
        margins: {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3},
        orientation: 'portrait',
        horizontalDpi: 4294967295,
        verticalDpi: 4294967295,
        fitToPage: !!(
          options.pageSetup &&
          (options.pageSetup.fitToWidth || options.pageSetup.fitToHeight) &&
          !options.pageSetup.scale
        ),
        pageOrder: 'downThenOver',
        blackAndWhite: false,
        draft: false,
        cellComments: 'None',
        errors: 'displayed',
        scale: 100,
        fitToWidth: 1,
        fitToHeight: 1,
        paperSize: undefined,
        showRowColHeaders: false,
        showGridLines: false,
        horizontalCentered: false,
        verticalCentered: false,
        rowBreaks: null,
        colBreaks: null,
      },
      options.pageSetup
    );

    // using shared strings creates a smaller xlsx file but may use more memory
    this.useSharedStrings = options.useSharedStrings || false;

    this._workbook = options.workbook;

    this.hasComments = false;

    // views
    this._views = options.views || [];

    // auto filter
    this.autoFilter = options.autoFilter || null;

    this._media = [];

    // start writing to stream now
    // TODO: [BP] does not wait for backpressure here
    this._writeOpenWorksheet();

    this.startedData = false;
  }

  get workbook() {
    return this._workbook;
  }

  get stream() {
    if (!this._stream) {
      // eslint-disable-next-line no-underscore-dangle
      this._stream = this._workbook._openStream(`/xl/worksheets/sheet${this.id}.xml`);
    }
    return this._stream;
  }

  // destroy - not a valid operation for a streaming writer
  // even though some streamers might be able to, it's a bad idea.
  destroy() {
    throw new Error('Invalid Operation: destroy');
  }

  async commit() {
    if (this.committed) {
      return;
    }
    // commit all rows
    for await (const cRow of this._rows) {
      if (cRow) {
        // write the row to the stream
        await this._writeRow(cRow);
      }
    }

    // we _cannot_ accept new rows from now on
    this._rows = null;

    if (!this.startedData) {
      await this._writeOpenSheetData();
    }
    await this._writeCloseSheetData();
    await this._writeAutoFilter();
    await this._writeMergeCells();

    // for some reason, Excel can't handle dimensions at the bottom of the file
    // await this._writeDimensions();

    await this._writeHyperlinks();
    await this._writeConditionalFormatting();
    await this._writeDataValidations();
    await this._writeSheetProtection();
    await this._writePageMargins();
    await this._writePageSetup();
    await this._writeBackground();

    // Legacy Data tag for comments
    await this._writeLegacyData();

    await this._writeCloseWorksheet();
    // signal end of stream to workbook
    this.stream.end();

    await this._sheetCommentsWriter.commit();
    // also commit the hyperlinks if any
    await this._sheetRelsWriter.commit();

    this.committed = true;
  }

  // return the current dimensions of the writer
  get dimensions() {
    return this._dimensions;
  }

  get views() {
    return this._views;
  }

  // =========================================================================
  // Columns

  // get the current columns array.
  get columns() {
    return this._columns;
  }

  // set the columns from an array of column definitions.
  // Note: any headers defined will overwrite existing values.
  set columns(value) {
    // calculate max header row count
    this._headerRowCount = value.reduce((pv, cv) => {
      const headerCount = (cv.header && 1) || (cv.headers && cv.headers.length) || 0;
      return Math.max(pv, headerCount);
    }, 0);

    // construct Column objects
    let count = 1;
    const columns = (this._columns = []);
    value.forEach(defn => {
      const column = new Column(this, count++, false);
      columns.push(column);
      column.defn = defn;
    });
  }

  getColumnKey(key) {
    return this._keys[key];
  }

  setColumnKey(key, value) {
    this._keys[key] = value;
  }

  deleteColumnKey(key) {
    delete this._keys[key];
  }

  eachColumnKey(f) {
    _.each(this._keys, f);
  }

  // get a single column by col number. If it doesn't exist, it and any gaps before it
  // are created.
  getColumn(c) {
    if (typeof c === 'string') {
      // if it matches a key'd column, return that
      const col = this._keys[c];
      if (col) return col;

      // otherwise, assume letter
      c = colCache.l2n(c);
    }
    if (!this._columns) {
      this._columns = [];
    }
    if (c > this._columns.length) {
      let n = this._columns.length + 1;
      while (n <= c) {
        this._columns.push(new Column(this, n++));
      }
    }
    return this._columns[c - 1];
  }

  // =========================================================================
  // Rows
  get _nextRow() {
    return this._rowZero + this._rows.length;
  }

  // iterate over every uncommitted row in the worksheet, including maybe empty rows
  eachRow(options, iteratee) {
    if (!iteratee) {
      iteratee = options;
      options = undefined;
    }
    if (options && options.includeEmpty) {
      const n = this._nextRow;
      for (let i = this._rowZero; i < n; i++) {
        iteratee(this.getRow(i), i);
      }
    } else {
      this._rows.forEach(row => {
        if (row.hasValues) {
          iteratee(row, row.number);
        }
      });
    }
  }

  async _commitRow(cRow) {
    // since rows must be written in order, we commit all rows up till and including cRow
    let found = false;
    while (this._rows.length && !found) {
      const row = this._rows.shift();
      this._rowZero++;
      if (row) {
        await this._writeRow(row);
        found = row.number === cRow.number;
        this._rowZero = row.number + 1;
      }
    }
  }

  get lastRow() {
    // returns last uncommitted row
    if (this._rows.length) {
      return this._rows[this._rows.length - 1];
    }
    return undefined;
  }

  // find a row (if exists) by row number
  findRow(rowNumber) {
    const index = rowNumber - this._rowZero;
    return this._rows[index];
  }

  getRow(rowNumber) {
    const index = rowNumber - this._rowZero;

    // may fail if rows have been comitted
    if (index < 0) {
      throw new Error('Out of bounds: this row has been committed');
    }
    let row = this._rows[index];
    if (!row) {
      this._rows[index] = row = new Row(this, rowNumber);
    }
    return row;
  }

  addRow(value) {
    const row = new Row(this, this._nextRow);
    this._rows[row.number - this._rowZero] = row;
    row.values = value;
    return row;
  }

  // ================================================================================
  // Cells

  // returns the cell at [r,c] or address given by r. If not found, return undefined
  findCell(r, c) {
    const address = colCache.getAddress(r, c);
    const row = this.findRow(address.row);
    return row ? row.findCell(address.column) : undefined;
  }

  // return the cell at [r,c] or address given by r. If not found, create a new one.
  getCell(r, c) {
    const address = colCache.getAddress(r, c);
    const row = this.getRow(address.row);
    return row.getCellEx(address);
  }

  mergeCells() {
    // may fail if rows have been comitted
    const dimensions = new Dimensions(Array.prototype.slice.call(arguments, 0)); // convert arguments into Array

    // check cells aren't already merged
    this._merges.forEach(merge => {
      if (merge.intersects(dimensions)) {
        throw new Error('Cannot merge already merged cells');
      }
    });

    // apply merge
    const master = this.getCell(dimensions.top, dimensions.left);
    for (let i = dimensions.top; i <= dimensions.bottom; i++) {
      for (let j = dimensions.left; j <= dimensions.right; j++) {
        if (i > dimensions.top || j > dimensions.left) {
          this.getCell(i, j).merge(master);
        }
      }
    }

    // index merge
    this._merges.push(dimensions);
  }

  // ===========================================================================
  // Conditional Formatting
  addConditionalFormatting(cf) {
    this.conditionalFormatting.push(cf);
  }

  removeConditionalFormatting(filter) {
    console.log('conditionalFormatting', this.conditionalFormatting);
    if (typeof filter === 'number') {
      this.conditionalFormatting.splice(filter, 1);
    } else if (filter instanceof Function) {
      this.conditionalFormatting = this.conditionalFormatting.filter(filter);
    } else {
      this.conditionalFormatting = [];
    }
  }

  // =========================================================================

  addBackgroundImage(imageId) {
    this._background = {
      imageId,
    };
  }

  getBackgroundImageId() {
    return this._background && this._background.imageId;
  }

  // ================================================================================

  async _write(buffer) {
    // Handle edge cases
    if (buffer instanceof StringBuf) {
      buffer = buffer.toBuffer();
    }

    if (!this.stream.write(buffer)) {
      // backpressure
      await once(this.stream, 'drain');
    }
  }

  _writeSheetProperties(xmlBuf, properties, pageSetup) {
    const sheetPropertiesModel = {
      outlineProperties: properties && properties.outlineProperties,
      tabColor: properties && properties.tabColor,
      pageSetup:
        pageSetup && pageSetup.fitToPage
          ? {
              fitToPage: pageSetup.fitToPage,
            }
          : undefined,
    };

    xmlBuf.addText(xform.sheetProperties.toXml(sheetPropertiesModel));
  }

  _writeSheetFormatProperties(xmlBuf, properties) {
    const sheetFormatPropertiesModel = properties
      ? {
          defaultRowHeight: properties.defaultRowHeight,
          dyDescent: properties.dyDescent,
          outlineLevelCol: properties.outlineLevelCol,
          outlineLevelRow: properties.outlineLevelRow,
        }
      : undefined;
    if (properties.defaultColWidth) {
      sheetFormatPropertiesModel.defaultColWidth = properties.defaultColWidth;
    }

    xmlBuf.addText(xform.sheetFormatProperties.toXml(sheetFormatPropertiesModel));
  }

  async _writeOpenWorksheet() {
    xmlBuffer.reset();

    xmlBuffer.addText('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    xmlBuffer.addText(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' mc:Ignorable="x14ac"' +
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
    );

    this._writeSheetProperties(xmlBuffer, this.properties, this.pageSetup);

    xmlBuffer.addText(xform.sheetViews.toXml(this.views));

    this._writeSheetFormatProperties(xmlBuffer, this.properties);

    await this._write(xmlBuffer);
  }

  async _writeColumns() {
    const cols = Column.toModel(this.columns);
    if (cols) {
      xform.columns.prepare(cols, {styles: this._workbook.styles});
      await this._write(xform.columns.toXml(cols));
    }
  }

  async _writeOpenSheetData() {
    await this._write('<sheetData>');
  }

  async _writeRow(row) {
    if (!this.startedData) {
      await this._writeColumns();
      await this._writeOpenSheetData();
      this.startedData = true;
    }

    if (row.hasValues || row.height) {
      const {model} = row;
      const options = {
        styles: this._workbook.styles,
        sharedStrings: this.useSharedStrings ? this._workbook.sharedStrings : undefined,
        hyperlinks: this._sheetRelsWriter.hyperlinksProxy,
        merges: this._merges,
        formulae: this._formulae,
        siFormulae: this._siFormulae,
        comments: [],
      };
      xform.row.prepare(model, options);
      await this._write(xform.row.toXml(model));

      if (options.comments.length) {
        this.hasComments = true;
        await this._sheetCommentsWriter.addComments(options.comments);
      }
    }
  }

  async _writeCloseSheetData() {
    await this._write('</sheetData>');
  }

  async _writeMergeCells() {
    if (this._merges.length) {
      xmlBuffer.reset();
      xmlBuffer.addText(`<mergeCells count="${this._merges.length}">`);
      this._merges.forEach(merge => {
        xmlBuffer.addText(`<mergeCell ref="${merge}"/>`);
      });
      xmlBuffer.addText('</mergeCells>');

      await this._write(xmlBuffer);
    }
  }

  async _writeHyperlinks() {
    // eslint-disable-next-line no-underscore-dangle
    await this._write(xform.hyperlinks.toXml(this._sheetRelsWriter._hyperlinks));
  }

  async _writeConditionalFormatting() {
    const options = {
      styles: this._workbook.styles,
    };
    xform.conditionalFormattings.prepare(this.conditionalFormatting, options);
    await this._write(xform.conditionalFormattings.toXml(this.conditionalFormatting));
  }

  async _writeDataValidations() {
    await this._write(xform.dataValidations.toXml(this.dataValidations.model));
  }

  async _writeSheetProtection() {
    await this._write(xform.sheetProtection.toXml(this.sheetProtection));
  }

  async _writePageMargins() {
    await this._write(xform.pageMargins.toXml(this.pageSetup.margins));
  }

  async _writePageSetup() {
    await this._write(xform.pageSeteup.toXml(this.pageSetup));
  }

  async _writeAutoFilter() {
    await this._write(xform.autoFilter.toXml(this.autoFilter));
  }

  async _writeBackground() {
    if (this._background) {
      if (this._background.imageId !== undefined) {
        const image = this._workbook.getImage(this._background.imageId);
        const pictureId = this._sheetRelsWriter.addMedia({
          Target: `../media/${image.name}`,
          Type: RelType.Image,
        });

        this._background = {
          ...this._background,
          rId: pictureId,
        };
      }
      await this._write(xform.picture.toXml({rId: this._background.rId}));
    }
  }

  async _writeLegacyData() {
    if (this.hasComments) {
      xmlBuffer.reset();
      xmlBuffer.addText(`<legacyDrawing r:id="${this._sheetCommentsWriter.vmlRelId}"/>`);
      await this._write(xmlBuffer);
    }
  }

  async _writeDimensions() {
    // for some reason, Excel can't handle dimensions at the bottom of the file
    // and we don't know the dimensions until the commit, so don't write them.
    // await this._write('<dimension ref="' + this._dimensions + '"/>');
  }

  async _writeCloseWorksheet() {
    await this._write('</worksheet>');
  }
}

module.exports = WorksheetWriter;
