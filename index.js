'use strict'
const RTFParser = require('./rtf-parser.js')
const RTFDocument = require('./rtf-document.js')
const RTFInterpreter = require('./rtf-interpreter.js')

module.exports = parse
parse.string = parseString
parse.stream = parseStream

function parseString (string, cb) {
  //parse(cb).end(string)
  var parser = parse(cb);
  var rtfParsed = parser.convert(string);
  
  const document = new RTFDocument()
  const interpreter = new RTFInterpreter(document)

  const errorHandler = err => {
      if (errored) return
      errored = true
      cb(err)
  };
  interpreter.setErrorHandler(errorHandler);

  for (var i = 0; i < rtfParsed.length; i++) {
      interpreter.write(rtfParsed[i], null);
  }

  cb(null, document);
}

function parseStream (stream, cb) {
  stream.pipe(parse(cb))
}

function parse (cb) {
  let errored = false
  const errorHandler = err => {
    if (errored) return
    errored = true
    //parser.unpipe(interpreter)
    //interpreter.end()
    cb(err)
  }
  //const document = new RTFDocument()
  const parser = new RTFParser()
  parser.once('error', errorHandler)
  //const interpreter = new RTFInterpreter(document)
  //interpreter.on('error', errorHandler)
  //interpreter.on('finish', () => {
  //  if (!errored) cb(null, document)
  //})
  //parser.pipe(interpreter)
  return parser
}

