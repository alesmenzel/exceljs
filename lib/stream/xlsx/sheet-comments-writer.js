const {once} = require('events');
const XmlStream = require('../../utils/xml-stream');
const RelType = require('../../xlsx/rel-type');
const colCache = require('../../utils/col-cache');
const CommentXform = require('../../xlsx/xform/comment/comment-xform');
const VmlNoteXform = require('../../xlsx/xform/comment/vml-note-xform');

class SheetCommentsWriter {
  constructor(worksheet, sheetRelsWriter, options) {
    // in a workbook, each sheet will have a number
    this.id = options.id;
    this.count = 0;
    this._worksheet = worksheet;
    this._workbook = options.workbook;
    this._sheetRelsWriter = sheetRelsWriter;
  }

  get commentsStream() {
    if (!this._commentsStream) {
      // eslint-disable-next-line no-underscore-dangle
      this._commentsStream = this._workbook._openStream(`/xl/comments${this.id}.xml`);
    }
    return this._commentsStream;
  }

  get vmlStream() {
    if (!this._vmlStream) {
      // eslint-disable-next-line no-underscore-dangle
      this._vmlStream = this._workbook._openStream(`xl/drawings/vmlDrawing${this.id}.vml`);
    }
    return this._vmlStream;
  }

  async _addRelationships() {
    const commentRel = {
      Type: RelType.Comments,
      Target: `../comments${this.id}.xml`,
    };

    const vmlDrawingRel = {
      Type: RelType.VmlDrawing,
      Target: `../drawings/vmlDrawing${this.id}.vml`,
    };

    const [vmlRelId] = await Promise.all([
      this._sheetRelsWriter.addRelationship(vmlDrawingRel),
      this._sheetRelsWriter.addRelationship(commentRel),
    ]);

    this.vmlRelId = vmlRelId;
  }

  _addCommentRefs() {
    this._workbook.commentRefs.push({
      commentName: `comments${this.id}`,
      vmlDrawing: `vmlDrawing${this.id}`,
    });
  }

  async _write(stream, data) {
    if (stream.write(data)) {
      // backpressure
      await once(stream, 'drain');
    }
  }

  async _writeOpen() {
    await Promise.all([
      this._write(
        this.commentsStream,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          '<authors><author>Author</author></authors>' +
          '<commentList>'
      ),
      this._write(
        this.vmlStream,
        '<?xml version="1.0" encoding="UTF-8"?>' +
          '<xml xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
          '<o:shapelayout v:ext="edit">' +
          '<o:idmap v:ext="edit" data="1" />' +
          '</o:shapelayout>' +
          '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
          '<v:stroke joinstyle="miter" />' +
          '<v:path gradientshapeok="t" o:connecttype="rect" />' +
          '</v:shapetype>'
      ),
    ]);
  }

  async _writeComment(comment, index) {
    const commentXform = new CommentXform();
    const commentsXmlStream = new XmlStream();
    commentXform.render(commentsXmlStream, comment);

    const vmlNoteXform = new VmlNoteXform();
    const vmlXmlStream = new XmlStream();
    vmlNoteXform.render(vmlXmlStream, comment, index);

    await Promise.all([
      this._write(this.commentsStream, commentsXmlStream.xml),
      this._write(this.vmlStream, vmlXmlStream.xml),
    ]);
  }

  async _writeClose() {
    await Promise.all([
      this._write(this.commentsStream, '</commentList></comments>'),
      this._write(this.vmlStream, '</xml>'),
    ]);
  }

  async addComments(comments) {
    if (comments && comments.length) {
      if (!this.startedData) {
        this._worksheet.comments = [];
        await this._writeOpen();
        await this._addRelationships();
        this._addCommentRefs();
        this.startedData = true;
      }

      comments.forEach(item => {
        item.refAddress = colCache.decodeAddress(item.ref);
      });

      for await (const comment of comments) {
        await this._writeComment(comment, this.count);
        this.count += 1;
      }
    }
  }

  async commit() {
    if (this.count) {
      await this._writeClose();
      this.commentsStream.end();
      this.vmlStream.end();
    }
  }
}

module.exports = SheetCommentsWriter;
