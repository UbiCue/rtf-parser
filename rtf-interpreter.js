'use strict'
const assert = require('assert')
const util = require('util')

const Writable = require('readable-stream').Writable
const RTFGroup = require('./rtf-group.js')
const RTFParagraph = require('./rtf-paragraph.js')
const RTFSpan = require('./rtf-span.js')
const iconv = require('iconv-lite')

const availableCP = [
  437, 737, 775, 850, 852, 853, 855, 857, 858, 860, 861, 863, 865, 866,
  869, 932, 1125, 1250, 1251, 1252, 1253, 1254, 1257 ]
const codeToCP = {
  0: 'ASCII',
  77: 'MacRoman',
  128: 'SHIFT_JIS',
  129: 'CP949', // Hangul
  130: 'JOHAB',
  134: 'CP936', // GB2312 simplified chinese
  136: 'BIG5',
  161: 'CP1253', // greek
  162: 'CP1254', // turkish
  163: 'CP1258', // vietnamese
  177: 'CP862', // hebrew
  178: 'CP1256', // arabic
  186: 'CP1257',  // baltic
  204: 'CP1251', // russian
  222: 'CP874', // thai
  238: 'CP238', // eastern european
  254: 'CP437' // PC-437
}

function RTFInterpreter(document) {
  //this = Writable({objectMode: true});
  Writable.call(this, {});
  this.objectMode = true;
  
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
  
  this.doc = document
  this.parserState = this.parseTop
  this.groupStack = []
  this.group = null
  this.once('prefinish', () => this.finisher())
  this.hexStore = []
  this.spanStyle = {}
  this._write = function(cmd, encoding, done) {
    const method = 'cmd$' + cmd.type.replace(/-(.)/g, (_, char) => char.toUpperCase())
    if (this[method]) {
      this[method](cmd)
    } else {
      process.emit('error', `Unknown RTF command ${cmd.type}, tried ${method}`)
    }
    done()
  }
  this.write = function(cmd, encoding) {
        const method = 'cmd$' + cmd.type.replace(/-(.)/g, (_, char) => char.toUpperCase());
        if (this[method]) {
            this[method](cmd)
        } else {
            //process.emit('error', `Unknown RTF command ${cmd.type}, tried ${method}`)
            console.log(`Unknown RTF command ${cmd.type}, tried ${method}`);
        }
    }
 
  var finisher = function() {
    while (this.groupStack.length) this.cmd$groupEnd()
    const initialStyle = this.doc.content.length ? this.doc.content[0].style : []
    for (let prop of Object.keys(this.doc.style)) {
      let match = true
      for (let para of this.doc.content) {
        if (initialStyle[prop] !== para.style[prop]) {
          match = false
          break
        }
      }
      if (match) this.doc.style[prop] = initialStyle[prop]
    }
  }
  var flushHexStore = function() {
    if (this.hexStore.length > 0) {
      let hexstr = this.hexStore.map(cmd => cmd.value).join('')
      this.group.addContent(new RTFSpan({
        value: iconv.decode(
          Buffer.from(hexstr, 'hex'), this.group.get('charset'))
      }))
      this.hexStore.splice(0)
    }
  }

  var cmd$groupStart = function() {
    this.flushHexStore()
    if (this.group) this.groupStack.push(this.group)
    this.group = new RTFGroup(this.group || this.doc)
  }
  var cmd$ignorable = function() {
    this.flushHexStore()
    this.group.ignorable = true
  }
  var cmd$endParagraph = function() {
    this.flushHexStore()
    this.group.addContent(new RTFParagraph())
  }
  var cmd$groupEnd = function() {
    this.flushHexStore()
    const endingGroup = this.group
    this.group = this.groupStack.pop()
    const doc = this.group || this.doc
    if (endingGroup instanceof FontTable) {
      doc.fonts = endingGroup.table
    } else if (endingGroup instanceof ColorTable) {
      doc.colors = endingGroup.table
    } else if (endingGroup !== this.doc && !endingGroup.get('ignorable')) {
      for (const item of endingGroup.content) {
        doc.documentAddContent(item)
      }
      process.emit('debug', 'GROUP END', endingGroup.type, endingGroup.get('ignorable'))
    }
  }
  var cmd$text = function(cmd) {
    this.flushHexStore()
    if (!this.group) { // an RTF fragment, missing the {\rtf1 header
      this.group = this.doc
    }
    //If there isn't already a style specified, use the current group style to start
    if (typeof cmd.style === 'undefined') {
        cmd.style = this.group.style;
    }
    //Update any styling specified for the current span
    cmd.style.bold = this.spanStyle.bold;
    cmd.style.italic = this.spanStyle.italic;
    this.group.addContent(new RTFSpan(cmd))
  }
  var cmd$controlWord = function(cmd) {
    this.flushHexStore()
    if (typeof this.group !== 'undefined' && this.group !== null) {
      if (!this.group.type) this.group.type = cmd.value
      const method = 'ctrl$' + cmd.value.replace(/-(.)/g, (_, char) => char.toUpperCase())
      if (this[method]) {
        this[method](cmd.param)
      } else {
        if (!this.group.get('ignorable')) process.emit('debug', method, cmd.param)
      }
    }
  }
  var cmd$hexchar = function(cmd) {
    this.hexStore.push(cmd)
  }
  var cmd$error = function(cmd) {
    this.emit('error', new Error('Error: ' + cmd.value + (cmd.row && cmd.col ? ' at line ' + cmd.row + ':' + cmd.col : '') + '.'))
  }

  var ctrl$rtf = function() {
    this.group = this.doc
  }

  // new line
  var ctrl$line = function() {
    this.group.addContent(new RTFSpan({ value: '\n' }))
  }

  // alignment
  var ctrl$qc = function() {
    this.group.style.align = 'center'
  }
  var ctrl$qj = function() {
    this.group.style.align = 'justify'
  }
  var ctrl$ql = function() {
    this.group.style.align = 'left'
  }
  var ctrl$qr = function() {
    this.group.style.align = 'right'
  }

  // text direction
  var ctrl$rtlch = function() {
    this.group.style.dir = 'rtl'
  }
  var ctrl$ltrch = function() {
    this.group.style.dir = 'ltr'
  }

  // general style
  var ctrl$par = function() {
    //Create new paragraph, starting from document styling
    this.group.addContent(new RTFParagraph(this.doc))
  }
  var ctrl$pard = function() {
    this.group.resetStyle()
  }
  var ctrl$plain = function() {
    this.group.style.fontSize = this.doc.getStyle('fontSize')
    //When explicitly setting to plain, set all styles to false for group and span
    this.group.style.bold = false;
    this.group.style.italic = false;
    this.group.style.underline = false;
    this.spanStyle.bold = false;
    this.spanStyle.italic = false;
  }
  var ctrl$b = function(set) {
    this.group.style.bold = set !== 0
    this.spanStyle.bold = this.group.style.bold
  }
  var ctrl$i = function(set) {
    this.group.style.italic = set !== 0
    this.spanStyle.italic = this.group.style.italic
  }
  var ctrl$u = function(num) {
    var charBuf = Buffer.alloc ? Buffer.alloc(2) : new Buffer(2)
    // RTF, for reasons, represents unicode characters as signed integers
    // thus managing to match literally no one.
    charBuf.writeInt16LE(num, 0)
    this.group.addContent(new RTFSpan({value: iconv.decode(charBuf, 'ucs2')}))
  }
  var ctrl$super = function() {
    this.group.style.valign = 'super'
  }
  var ctrl$sub = function() {
    this.group.style.valign = 'sub'
  }
  var ctrl$nosupersub = function() {
    this.group.style.valign = 'normal'
  }
  var ctrl$strike = function(set) {
    this.group.style.strikethrough = set !== 0
  }
  var ctrl$ul = function(set) {
    this.group.style.underline = set !== 0
  }
  var ctrl$ulnone = function(set) {
    this.group.style.underline = false
  }
  var ctrl$fi = function(value) {
    this.group.style.firstLineIndent = value
  }
  var ctrl$cufi = function(value) {
    this.group.style.firstLineIndent = value * 100
  }
  var ctrl$li = function(value) {
    this.group.style.indent = value
  }
  var ctrl$lin = function(value) {
    this.group.style.indent = value
  }
  var ctrl$culi = function(value) {
    this.group.style.indent = value * 100
  }
  var ctrl$tab = function() {
      var spacer = { value: "&nbsp;", style: this.group.style };
      this.group.addContent(new RTFSpan(spacer));
  }

// encodings
  var ctrl$ansi = function() {
    this.group.charset = 'ASCII'
  }
  var ctrl$mac = function() {
    this.group.charset = 'MacRoman'
  }
  var ctrl$pc = function() {
    this.group.charset = 'CP437'
  }
  var ctrl$pca = function() {
    this.group.charset = 'CP850'
  }
  var ctrl$ansicpg = function(codepage) {
    if (availableCP.indexOf(codepage) === -1) {
      this.emit('error', new Error('Codepage ' + codepage + ' is not available.'))
    } else {
      this.group.charset = 'CP' + codepage
    }
  }

// fonts
  var ctrl$fonttbl = function() {
    this.group = new FontTable(this.group.parent)
  }
  var ctrl$f = function(num) {
    if (this.group instanceof FontTable) {
      this.group.currentFont = this.group.table[num] = new Font()
    } else if (this.group.parent instanceof FontTable) {
      this.group.parent.currentFont = this.group.parent.table[num] = new Font()
    } else {
      this.group.style.font = num
    }
  }
  var ctrl$fnil = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'nil'
    }
  }
  var ctrl$froman = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'roman'
    }
  }
  var ctrl$fswiss = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'swiss'
    }
  }
  var ctrl$fmodern = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'modern'
    }
  }
  var ctrl$fscript = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'script'
    }
  }
  var ctrl$fdecor = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'decor'
    }
  }
  var ctrl$ftech = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'tech'
    }
  }
  var ctrl$fbidi = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'bidi'
    }
  }
  var ctrl$fcharset = function(code) {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      let charset = null
      if (code === 1) {
        charset = this.group.get('charset')
      } else {
        charset = codeToCP[code]
      }
      if (charset == null) {
        return this.emit('error', new Error('Unsupported charset code #' + code))
      }
      this.group.get('currentFont').charset = charset
    }
  }
  var ctrl$fprq = function(pitch) {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').pitch = pitch
    }
  }

  // colors
  var ctrl$colortbl = function() {
    this.group = new ColorTable(this.group.parent)
  }
  var ctrl$red = function(value) {
    if (this.group instanceof ColorTable) {
      this.group.red = value
    }
  }
  var ctrl$blue = function(value) {
    if (this.group instanceof ColorTable) {
      this.group.blue = value
    }
  }
  var ctrl$green = function(value) {
    if (this.group instanceof ColorTable) {
      this.group.green = value
    }
  }
  var ctrl$cf = function(value) {
    this.group.style.foreground = value
  }
  var ctrl$cb = function(value) {
    this.group.style.background = value
  }
  var ctrl$fs = function(value) {
    this.group.style.fontSize = value
  }

// margins
  var ctrl$margl = function(value) {
    this.doc.marginLeft = value
  }
  var ctrl$margr = function(value) {
    this.doc.marginRight = value
  }
  var ctrl$margt = function(value) {
    this.doc.marginTop = value
  }
  var ctrl$margb = function(value) {
    this.doc.marginBottom = value
  }

// unsupported (and we need to ignore content)
  var ctrl$stylesheet = function(value) {
    this.group.ignorable = true
  }
  var ctrl$info = function(value) {
    this.group.ignorable = true
  }
  var ctrl$mmathPr = function(value) {
    this.group.ignorable = true
  }
}

function FontTable(parent) {
  //this = RTFGroup(parent);
  RTFGroup.call(this, parent);
  
  this.table = []
  this.currentFont = {family: 'roman', charset: 'ASCII', name: 'Serif'}
  this.addContent = function(text) {
    this.currentFont.name += text.value.replace(/;\s*$/, '')
  }
}

function Font() {
  this.family = null
  this.charset = null
  this.name = ''
  this.pitch = 0
}

function ColorTable(parent) {
  //this = RTFGroup(parent);
  RTFGroup.call(this, parent);
  
  this.table = []
  this.red = 0
  this.blue = 0
  this.green = 0
  this.addContent = function(text) {
    assert(text.value === ';', 'got: ' + util.inspect(text))
    this.table.push({
      red: this.red,
      blue: this.blue,
      green: this.green
    })
    this.red = 0
    this.blue = 0
    this.green = 0
  }
}

module.exports = RTFInterpreter
