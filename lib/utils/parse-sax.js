const {SaxesParser} = require('saxes');
const {PassThrough} = require('stream');
const {bufferToString} = require('./browser-buffer-decode');

// Namespaces the xform parsers expect to be the default (unprefixed) namespace.
// OOXML allows binding them to a prefix instead — .NET OpenXML SDK / ClosedXML emit
// '<x:workbook xmlns:x="...">' and Hancell additionally emits '<ep:Properties ...>'
// in docProps/app.xml. Such files are valid (Excel opens them) but parse to empty
// models here. Detect the binding and strip that prefix from tag names so prefixed
// and unprefixed documents parse identically.
const DEFAULT_NS_LIST = [
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
];
const XMLNS_PREFIX_PATTERN = /^xmlns:(\w+)$/;

function findDefaultNsPrefix(attributes) {
  for (const [name, value] of Object.entries(attributes)) {
    if (DEFAULT_NS_LIST.includes(value)) {
      const match = name.match(XMLNS_PREFIX_PATTERN);
      if (match) {
        return `${match[1]}:`;
      }
    }
  }
  return null;
}

module.exports = async function* (iterable) {
  // TODO: Remove once node v8 is deprecated
  // Detect and upgrade old streams
  if (iterable.pipe && !iterable[Symbol.asyncIterator]) {
    iterable = iterable.pipe(new PassThrough());
  }
  const saxesParser = new SaxesParser();
  let error;
  saxesParser.on('error', err => {
    error = err;
  });
  let events = [];
  let nsPrefix = null;
  // saxes keeps the emitted tag objects on its internal stack for well-formedness
  // checks, so return a copy instead of mutating the original.
  const stripPrefix = value => {
    if (nsPrefix && value.name.startsWith(nsPrefix)) {
      return {...value, name: value.name.substring(nsPrefix.length)};
    }
    return value;
  };
  saxesParser.on('opentag', value => {
    if (nsPrefix === null && value.name.includes(':')) {
      nsPrefix = findDefaultNsPrefix(value.attributes) || '';
    }
    events.push({eventType: 'opentag', value: stripPrefix(value)});
  });
  saxesParser.on('text', value => events.push({eventType: 'text', value}));
  saxesParser.on('closetag', value => events.push({eventType: 'closetag', value: stripPrefix(value)}));
  for await (const chunk of iterable) {
    saxesParser.write(bufferToString(chunk));
    // saxesParser.write and saxesParser.on() are synchronous,
    // so we can only reach the below line once all events have been emitted
    if (error) throw error;
    // As a performance optimization, we gather all events instead of passing
    // them one by one, which would cause each event to go through the event queue
    yield events;
    events = [];
  }
};
