/* eslint-disable max-classes-per-file */
const {once} = require('events');
const utils = require('../../utils/utils');
const RelType = require('../../xlsx/rel-type');

class HyperlinksProxy {
  constructor(sheetRelsWriter) {
    this.writer = sheetRelsWriter;
  }

  async push(hyperlink) {
    await this.writer.addHyperlink(hyperlink);
  }
}

class SheetRelsWriter {
  constructor(options) {
    // in a workbook, each sheet will have a number
    this.id = options.id;

    // count of all relationships
    this.count = 0;

    // keep record of all hyperlinks
    this._hyperlinks = [];

    this._workbook = options.workbook;
  }

  get stream() {
    if (!this._stream) {
      // eslint-disable-next-line no-underscore-dangle
      this._stream = this._workbook._openStream(`/xl/worksheets/_rels/sheet${this.id}.xml.rels`);
    }
    return this._stream;
  }

  async _write(data) {
    if (!this.stream.push(data)) {
      await once(this.stream, 'drain');
    }
  }

  get length() {
    return this._hyperlinks.length;
  }

  each(fn) {
    return this._hyperlinks.forEach(fn);
  }

  get hyperlinksProxy() {
    return this._hyperlinksProxy || (this._hyperlinksProxy = new HyperlinksProxy(this));
  }

  async addHyperlink(hyperlink) {
    // Write to stream
    const relationship = {
      Target: hyperlink.target,
      Type: RelType.Hyperlink,
      TargetMode: 'External',
    };
    const rId = await this._writeRelationship(relationship);

    // store sheet stuff for later
    this._hyperlinks.push({
      rId,
      address: hyperlink.address,
    });
  }

  async addMedia(media) {
    return this._writeRelationship(media);
  }

  async addRelationship(rel) {
    return this._writeRelationship(rel);
  }

  async commit() {
    if (this.count) {
      // write xml utro
      await this._writeClose();
      // and close stream
      this.stream.end();
    }
  }

  // ================================================================================
  async _writeOpen() {
    await this._write(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    );
  }

  async _writeRelationship(relationship) {
    if (!this.count) {
      await this._writeOpen();
    }

    const rId = `rId${++this.count}`;

    if (relationship.TargetMode) {
      await this._write(
        `<Relationship Id="${rId}"` +
          ` Type="${relationship.Type}"` +
          ` Target="${utils.xmlEncode(relationship.Target)}"` +
          ` TargetMode="${relationship.TargetMode}"` +
          '/>'
      );
    } else {
      await this._write(
        `<Relationship Id="${rId}" Type="${relationship.Type}" Target="${relationship.Target}"/>`
      );
    }

    return rId;
  }

  async _writeClose() {
    await this._write('</Relationships>');
  }
}

module.exports = SheetRelsWriter;
