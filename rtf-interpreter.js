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
  this.once('prefinish', function () { this.finisher(); })
  this.hexStore = []
  this.spanStyle = {}
  this.deriveCmd = function(type, prefix) {
	let command = '';
		for (var i=0; i<type.length; i++) {
			if ((type[i] == '-') && (i<type.length+1)) {
				command += type[i+1].toUpperCase();
				i++;
			}
			else {
				command += type[i];
			}
		}
		//const method = 'cmd$' + cmd.type.replace(/-(.)/g, (_, char) => char.toUpperCase());
		return prefix+'$'+command;
	}
	this._write = function(cmd, encoding, done) {
		let method = this.deriveCmd(cmd.type, 'cmd');
		if (this[method]) {
		  this[method](cmd)
		} else {
		  process.emit('error', "Unknown RTF command "+cmd.type+", tried "+method)
		}
		done()
	}
	this.write = function(cmd, encoding) {
		//const method = 'cmd$' + cmd.type.replace(/-(.)/g, (_, char) => char.toUpperCase());
		let method = this.deriveCmd(cmd.type, 'cmd');
		if (this[method]) {
			this[method](cmd)
		} else {
			//process.emit('error', `Unknown RTF command ${cmd.type}, tried ${method}`)
			console.log("Unknown RTF command "+cmd.type+", tried "+method);
		}
	}
 
	this.finisher = function() {
		while (this.groupStack.length) this.cmd$groupEnd()
		const initialStyle = this.doc.content.length ? this.doc.content[0].style : []
		for (let i=0; i<Object.keys(this.doc.style).length;i++) {
			let prop = Object.keys(this.doc.style)[i];
			let match = true
			for (let j=0; j<this.doc.content.length; j++) {
				let para = this.doc.content[j];
				if (initialStyle[prop] !== para.style[prop]) {
				  match = false
				  break
				}
			}
			if (match) this.doc.style[prop] = initialStyle[prop]
		}
	}
  this.flushHexStore = function() {
    if (this.hexStore.length > 0) {
      let hexstr = this.hexStore.map(function(cmd) { cmd.value }).join('')
      this.group.addContent(new RTFSpan({
        value: iconv.decode(
          Buffer.from(hexstr, 'hex'), this.group.get('charset'))
      }))
      this.hexStore.splice(0)
    }
  }

  this.cmd$groupStart = function() {
    this.flushHexStore()
    if (this.group) this.groupStack.push(this.group)
    this.group = new RTFGroup(this.group || this.doc)
  }
  this.cmd$ignorable = function() {
    this.flushHexStore()
    this.group.ignorable = true
  }
  this.cmd$endParagraph = function() {
    this.flushHexStore()
    this.group.addContent(new RTFParagraph())
  }
  this.cmd$groupEnd = function() {
	this.flushHexStore()
	const endingGroup = this.group
	this.group = this.groupStack.pop()
	const doc = this.group || this.doc
	if (endingGroup instanceof FontTable) {
		doc.fonts = endingGroup.table
	} else if (endingGroup instanceof ColorTable) {
		doc.colors = endingGroup.table
	} else if (endingGroup !== this.doc && !endingGroup.get('ignorable')) {
		for (var i=0; i<endingGroup.content.length; i++) {
			var item = endingGroup.content[i];
			this.addContent(doc, item);
		}
		process.emit('debug', 'GROUP END', endingGroup.type, endingGroup.get('ignorable'))
	}
  }
  this.addContent = function(destination, content) {
	  if (typeof destination.docAddContent !== 'undefined') {
			destination.docAddContent(content)
		}
		else if (typeof destination.addContent !== 'undefined') {
			destination.addContent(content);
		}
  }
  this.cmd$text = function(cmd) {
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
  this.cmd$controlWord = function(cmd) {
    this.flushHexStore()
    if (typeof this.group !== 'undefined' && this.group !== null) {
      if (!this.group.type) this.group.type = cmd.value;
      	const method = this.deriveCmd(cmd.value, 'ctrl');
      	if (this[method]) {
        	this[method](cmd.param)
      	} else {
        	if (!this.group.get('ignorable')) process.emit('debug', method, cmd.param)
      	}
      }
  }
  this.cmd$hexchar = function(cmd) {
    this.hexStore.push(cmd)
  }
  this.cmd$error = function(cmd) {
    console.log('Error: ' + cmd.value + (cmd.row && cmd.col ? ' at line ' + cmd.row + ':' + cmd.col : '') + '.');
    //this.emit('error', new Error('Error: ' + cmd.value + (cmd.row && cmd.col ? ' at line ' + cmd.row + ':' + cmd.col : '') + '.'))
  }

  this.ctrl$rtf = function() {
    //this.group = this.doc
  }

  // new line
  this.ctrl$line = function() {
    this.group.addContent(new RTFSpan({ value: '\n' }))
  }

  // alignment
  this.ctrl$qc = function() {
    this.group.style.align = 'center'
  }
  this.ctrl$qj = function() {
    this.group.style.align = 'justify'
  }
  this.ctrl$ql = function() {
    this.group.style.align = 'left'
  }
  this.ctrl$qr = function() {
    this.group.style.align = 'right'
  }

  // text direction
  this.ctrl$rtlch = function() {
    this.group.style.dir = 'rtl'
  }
  this.ctrl$ltrch = function() {
    this.group.style.dir = 'ltr'
  }

  // general style
  this.ctrl$par = function() {
    //Create new paragraph, starting from document styling
    this.group.addContent(new RTFParagraph(this.doc))
  }
  this.ctrl$pard = function() {
    this.group.resetStyle()
  }
  this.ctrl$plain = function() {
    this.group.style.fontSize = this.doc.getStyle('fontSize')
    //When explicitly setting to plain, set all styles to false for group and span
    this.group.style.bold = false;
    this.group.style.italic = false;
    this.group.style.underline = false;
    this.spanStyle.bold = false;
    this.spanStyle.italic = false;
  }
  this.ctrl$b = function(set) {
    this.group.style.bold = set !== 0
    this.spanStyle.bold = this.group.style.bold
  }
  this.ctrl$i = function(set) {
    this.group.style.italic = set !== 0
    this.spanStyle.italic = this.group.style.italic
  }
  this.ctrl$u = function(num) {
    var charBuf = Buffer.alloc ? Buffer.alloc(2) : new Buffer(2)
    // RTF, for reasons, represents unicode characters as signed integers
    // thus managing to match literally no one.
    charBuf.writeInt16LE(num, 0)
    this.group.addContent(new RTFSpan({value: iconv.decode(charBuf, 'ucs2')}))
  }
  this.ctrl$super = function() {
    this.group.style.valign = 'super'
  }
  this.ctrl$sub = function() {
    this.group.style.valign = 'sub'
  }
  this.ctrl$nosupersub = function() {
    this.group.style.valign = 'normal'
  }
  this.ctrl$strike = function(set) {
    this.group.style.strikethrough = set !== 0
  }
  this.ctrl$ul = function(set) {
    this.group.style.underline = set !== 0
  }
  this.ctrl$ulnone = function(set) {
    this.group.style.underline = false
  }
  this.ctrl$fi = function(value) {
    this.group.style.firstLineIndent = value
  }
  this.ctrl$cufi = function(value) {
    this.group.style.firstLineIndent = value * 100
  }
  this.ctrl$li = function(value) {
    this.group.style.indent = value
  }
  this.ctrl$lin = function(value) {
    this.group.style.indent = value
  }
  this.ctrl$culi = function(value) {
    this.group.style.indent = value * 100
  }
  this.ctrl$tab = function() {
      var spacer = { value: "&nbsp;", style: this.group.style };
      this.group.addContent(new RTFSpan(spacer));
  }

// encodings
  this.ctrl$ansi = function() {
    this.group.charset = 'ASCII'
  }
  this.ctrl$mac = function() {
    this.group.charset = 'MacRoman'
  }
  this.ctrl$pc = function() {
    this.group.charset = 'CP437'
  }
  this.ctrl$pca = function() {
    this.group.charset = 'CP850'
  }
  this.ctrl$ansicpg = function(codepage) {
    if (availableCP.indexOf(codepage) === -1) {
    	console.log('Codepage ' + codepage + ' is not available.');
      	//this.emit('error', new Error('Codepage ' + codepage + ' is not available.'))
    } else {
      	this.group.charset = 'CP' + codepage
    }
  }

// fonts
  this.ctrl$fonttbl = function() {
    this.group = new FontTable(this.group.parent)
  }
  this.ctrl$f = function(num) {
    if (this.group instanceof FontTable) {
      this.group.currentFont = this.group.table[num] = new Font()
    } else if (this.group.parent instanceof FontTable) {
      this.group.parent.currentFont = this.group.parent.table[num] = new Font()
    } else {
      this.group.style.font = num
    }
  }
  this.ctrl$fnil = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'nil'
    }
  }
  this.ctrl$froman = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'roman'
    }
  }
  this.ctrl$fswiss = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'swiss'
    }
  }
  this.ctrl$fmodern = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'modern'
    }
  }
  this.ctrl$fscript = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'script'
    }
  }
  this.ctrl$fdecor = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'decor'
    }
  }
  this.ctrl$ftech = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'tech'
    }
  }
  this.ctrl$fbidi = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'bidi'
    }
  }
  this.ctrl$fcharset = function(code) {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      let charset = null
      if (code === 1) {
        charset = this.group.get('charset')
      } else {
        charset = codeToCP[code]
      }
      if (charset == null) {
      	console.log('Unsupported charset code #' + code);
        //return this.emit('error', new Error('Unsupported charset code #' + code))
      }
      this.group.get('currentFont').charset = charset
    }
  }
  this.ctrl$fprq = function(pitch) {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').pitch = pitch
    }
  }

  // colors
  this.ctrl$colortbl = function() {
    this.group = new ColorTable(this.group.parent)
  }
  this.ctrl$red = function(value) {
    if (this.group instanceof ColorTable) {
      this.group.red = value
    }
  }
  this.ctrl$blue = function(value) {
    if (this.group instanceof ColorTable) {
      this.group.blue = value
    }
  }
  this.ctrl$green = function(value) {
    if (this.group instanceof ColorTable) {
      this.group.green = value
    }
  }
  this.ctrl$cf = function(value) {
    this.group.style.foreground = value
  }
  this.ctrl$cb = function(value) {
    this.group.style.background = value
  }
  this.ctrl$fs = function(value) {
    this.group.style.fontSize = value
  }

// margins
  this.ctrl$margl = function(value) {
    this.doc.marginLeft = value
  }
  this.ctrl$margr = function(value) {
    this.doc.marginRight = value
  }
  this.ctrl$margt = function(value) {
    this.doc.marginTop = value
  }
  this.ctrl$margb = function(value) {
    this.doc.marginBottom = value
  }

// unsupported (and we need to ignore content)
  this.ctrl$stylesheet = function(value) {
    this.group.ignorable = true
  }
  this.ctrl$info = function(value) {
    this.group.ignorable = true
  }
  this.ctrl$mmathPr = function(value) {
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

// Production steps of ECMA-262, Edition 5, 15.4.4.19
// Reference: http://es5.github.io/#x15.4.4.19
if (!Array.prototype.map) {

  Array.prototype.map = function(callback/*, thisArg*/) {

    var T, A, k;

    if (this == null) {
      throw new TypeError('this is null or not defined');
    }

    // 1. Let O be the result of calling ToObject passing the |this| 
    //    value as the argument.
    var O = Object(this);

    // 2. Let lenValue be the result of calling the Get internal 
    //    method of O with the argument "length".
    // 3. Let len be ToUint32(lenValue).
    var len = O.length >>> 0;

    // 4. If IsCallable(callback) is false, throw a TypeError exception.
    // See: http://es5.github.com/#x9.11
    if (typeof callback !== 'function') {
      throw new TypeError(callback + ' is not a function');
    }

    // 5. If thisArg was supplied, let T be thisArg; else let T be undefined.
    if (arguments.length > 1) {
      T = arguments[1];
    }

    // 6. Let A be a new array created as if by the expression new Array(len) 
    //    where Array is the standard built-in constructor with that name and 
    //    len is the value of len.
    A = new Array(len);

    // 7. Let k be 0
    k = 0;

    // 8. Repeat, while k < len
    while (k < len) {

      var kValue, mappedValue;

      // a. Let Pk be ToString(k).
      //   This is implicit for LHS operands of the in operator
      // b. Let kPresent be the result of calling the HasProperty internal 
      //    method of O with argument Pk.
      //   This step can be combined with c
      // c. If kPresent is true, then
      if (k in O) {

        // i. Let kValue be the result of calling the Get internal 
        //    method of O with argument Pk.
        kValue = O[k];

        // ii. Let mappedValue be the result of calling the Call internal 
        //     method of callback with T as the this value and argument 
        //     list containing kValue, k, and O.
        mappedValue = callback.call(T, kValue, k, O);

        // iii. Call the DefineOwnProperty internal method of A with arguments
        // Pk, Property Descriptor
        // { Value: mappedValue,
        //   Writable: true,
        //   Enumerable: true,
        //   Configurable: true },
        // and false.

        // In browsers that support Object.defineProperty, use the following:
        // Object.defineProperty(A, k, {
        //   value: mappedValue,
        //   writable: true,
        //   enumerable: true,
        //   configurable: true
        // });

        // For best browser support, use the following:
        A[k] = mappedValue;
      }
      // d. Increase k by 1.
      k++;
    }

    // 9. return A
    return A;
  };
}

module.exports = RTFInterpreter
