'use strict'
const Transform = require('readable-stream').Transform

function RTFParser() {
  //this = Transform({objectMode: true});
  Transform.call(this, {});
  this.objectMode = true;
  this.text = ''
  this.controlWord = ''
  this.controlWordParam = ''
  this.hexChar = ''
  this.parserState = this.parseText
  this.char = 0
  this.row = 1
  this.col = 1
  this._transform = function(buf, encoding, done) {
    const text = buf.toString('ascii')
    for (let ii = 0; ii < text.length; ++ii) {
      ++this.char
      if (text[ii] === '\n') {
        ++this.row
        this.col = 1
      } else {
        ++this.col
      }
      this.parserState(text[ii])
    }
    done()
  }
  this._flush = function(done) {
    if (this.text !== '\u0000') this.emitText()
    done()
  }
  this.convert(text) {
        for (let ii = 0; ii < text.length; ++ii) {
            ++this.char
            if (text[ii] === '\n') {
                ++this.row
                this.col = 1
            } else {
                ++this.col
            }
            this.parserState(text[ii])
        }

        return (this.fullText);
    }
  var parseText = function(char) {
    if (char === '\\') {
      this.parserState = this.parseEscapes
    } else if (char === '{') {
      this.emitStartGroup()
    } else if (char === '}') {
      this.emitEndGroup()
    } else if (char === '\x0A' || char === '\x0D') {
      // cr/lf are noise chars
    } else {
      this.text += char
    }
  }

  var parseEscapes = function(char) {
    if (char === '\\' || char === '{' || char === '}') {
      this.text += char
      this.parserState = this.parseText
    } else {
      this.parserState = this.parseControlSymbol
      this.parseControlSymbol(char)
    }
  }
  var parseControlSymbol = function(char) {
    if (char === '~') {
      this.text += '\u00a0' // nbsp
      this.parserState = this.parseText
    } else if (char === '-') {
      this.text += '\u00ad' // soft hyphen
    } else if (char === '_') {
      this.text += '\u2011' // non-breaking hyphen
    } else if (char === '*') {
      this.emitIgnorable()
      this.parserState = this.parseText
    } else if (char === "'") {
      this.parserState = this.parseHexChar
    } else if (char === '|') { // formula character
      this.emitFormula()
      this.parserState = this.parseText
    } else if (char === ':') { // subentry in an index entry
      this.emitIndexSubEntry()
      this.parserState = this.parseText
    } else if (char === '\x0a') {
      this.emitEndParagraph()
      this.parserState = this.parseText
    } else if (char === '\x0d') {
      this.emitEndParagraph()
      this.parserState = this.parseText
    } else {
      this.parserState = this.parseControlWord
      this.parseControlWord(char)
    }
  }
  var parseHexChar = function(char) {
    if (/^[A-Fa-f0-9]$/.test(char)) {
      this.hexChar += char
      if (this.hexChar.length >= 2) {
        this.emitHexChar()
        this.parserState = this.parseText
      }
    } else {
      this.emitError(`Invalid character "${char}" in hex literal.`)
      this.parserState = this.parseText
    }
  }
  var parseControlWord = function(char) {
    if (char === ' ') {
      this.emitControlWord()
      this.parserState = this.parseText
    } else if (/^[-\d]$/.test(char)) {
      this.parserState = this.parseControlWordParam
      this.controlWordParam += char
    } else if (/^[A-Za-z]$/.test(char)) {
      this.controlWord += char
    } else {
      this.emitControlWord()
      this.parserState = this.parseText
      this.parseText(char)
    }
  }
  var parseControlWordParam = function(char) {
    if (/^\d$/.test(char)) {
      this.controlWordParam += char
    } else if (char === ' ') {
      this.emitControlWord()
      this.parserState = this.parseText
    } else {
      this.emitControlWord()
      this.parserState = this.parseText
      this.parseText(char)
    }
  }
  var emitText = function() {
    if (this.text === '') return
    this.push({type: 'text', value: this.text, pos: this.char, row: this.row, col: this.col})
    this.text = ''
  }
  var emitControlWord = function() {
    this.emitText()
    if (this.controlWord === '') {
      this.emitError('empty control word')
    } else {
      this.push({
        type: 'control-word',
        value: this.controlWord,
        param: this.controlWordParam !== '' && Number(this.controlWordParam),
        pos: this.char,
        row: this.row,
        col: this.col
      })
    }
    this.controlWord = ''
    this.controlWordParam = ''
  }
  var emitStartGroup = function() {
    this.emitText()
    this.push({type: 'group-start', pos: this.char, row: this.row, col: this.col})
  }
  var emitEndGroup = function() {
    this.emitText()
    this.push({type: 'group-end', pos: this.char, row: this.row, col: this.col})
  }
  var emitIgnorable = function() {
    this.emitText()
    this.push({type: 'ignorable', pos: this.char, row: this.row, col: this.col})
  }
  var emitHexChar = function() {
    this.emitText()
    this.push({type: 'hexchar', value: this.hexChar, pos: this.char, row: this.row, col: this.col})
    this.hexChar = ''
  }
  var emitError = function(message) {
    this.emitText()
    this.push({type: 'error', value: message, row: this.row, col: this.col, char: this.char, stack: new Error().stack})
  }
  var emitEndParagraph = function() {
    this.emitText()
    this.push({type: 'end-paragraph', pos: this.char, row: this.row, col: this.col})
  }
  
  //Explicitly define once
  this.once = function(context, fn) {
    var result;

    return function () {
        if (fn) {
            result = fn.apply(context || this, arguments);
            fn = null;
        }

        return result;
    }
  }
}

module.exports = RTFParser
