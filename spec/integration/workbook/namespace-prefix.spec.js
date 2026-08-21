const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const SPREADSHEET_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const EXTENDED_PROPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';

function prefixDefaultNs(xml, ns, prefix) {
  return xml
    .replace(new RegExp(`xmlns="${ns}"`, 'g'), `xmlns:${prefix}="${ns}"`)
    .replace(/<(\/?)(?!\?|!)([a-zA-Z][^\s/>:]*)(?=[\s/>])/g, `<$1${prefix}:$2`);
}

// .NET OpenXML SDK / ClosedXML emit spreadsheetml elements with a namespace prefix
// ('<x:workbook xmlns:x="...">'). Rewrite a workbook produced by exceljs into that
// dialect to verify both the document and streaming parsers accept it.
async function buildPrefixedXlsxBuffer() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('corpus');
  ws.addRow(['id', 'source', 'target']);
  ws.addRow([1, 'hello', 'bonjour']);
  ws.addRow([2, 'world', 'monde']);
  const buffer = await wb.xlsx.writeBuffer();

  const zip = await JSZip.loadAsync(buffer);
  // Rebuild with xl/workbook.xml ahead of the worksheets — the entry order .NET/POI
  // writers use. (exceljs writes workbook.xml last, which trips the unrelated
  // stream-reader ordering bug exceljs#3064 and would make this spec flaky.)
  const rebuilt = new JSZip();
  const rank = name => {
    if (name === 'xl/workbook.xml') return 0;
    if (name.startsWith('xl/worksheets/')) return 2;
    return 1;
  };
  const names = Object.keys(zip.files)
    .filter(name => !zip.files[name].dir)
    .sort((a, b) => rank(a) - rank(b));
  const contents = await Promise.all(names.map(name => zip.files[name].async('string')));
  names.forEach((name, i) => {
    let xml = contents[i];
    if (/^xl\/.*\.xml$/.test(name) && xml.includes(`xmlns="${SPREADSHEET_MAIN_NS}"`)) {
      xml = prefixDefaultNs(xml, SPREADSHEET_MAIN_NS, 'x');
    } else if (name === 'docProps/app.xml' && xml.includes(`xmlns="${EXTENDED_PROPS_NS}"`)) {
      // Hancell binds the extended-properties namespace to a prefix as well
      xml = prefixDefaultNs(xml, EXTENDED_PROPS_NS, 'ep');
    }
    rebuilt.file(name, xml);
  });
  return rebuilt.generateAsync({type: 'nodebuffer'});
}

describe('Workbook', () => {
  describe('Namespace-prefixed OOXML (OpenXML SDK / ClosedXML dialect)', () => {
    it('reads a prefixed workbook with xlsx.load', async () => {
      const buffer = await buildPrefixedXlsxBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      expect(wb.worksheets).to.have.length(1);
      expect(wb.worksheets[0].name).to.equal('corpus');
      expect(wb.worksheets[0].getCell('B2').value).to.equal('hello');
      expect(wb.worksheets[0].getCell('C3').value).to.equal('monde');
    });

    it('reads a prefixed workbook with the streaming WorkbookReader', async () => {
      const buffer = await buildPrefixedXlsxBuffer();
      const {PassThrough} = require('stream');
      const input = new PassThrough();
      input.end(buffer);
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(input, {});
      const rows = [];
      for await (const worksheet of reader) {
        for await (const row of worksheet) {
          rows.push(row.values.slice(1));
        }
      }
      expect(rows).to.deep.equal([
        ['id', 'source', 'target'],
        [1, 'hello', 'bonjour'],
        [2, 'world', 'monde'],
      ]);
    });

    it('still reads unprefixed workbooks identically', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('plain');
      ws.addRow(['a', 'b']);
      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      expect(wb2.worksheets[0].name).to.equal('plain');
      expect(wb2.worksheets[0].getCell('A1').value).to.equal('a');
    });
  });
});
