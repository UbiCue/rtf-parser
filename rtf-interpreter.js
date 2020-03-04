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
  2: 'SYMBOL',
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
  this.inTxField = false;
  this.abandonDocument = false;
  this.savedCharset = '';
  
  this.defaultTabWidth = 720;
  this.tabWidth = this.defaultTabWidth;
  this.spacing = 0; //0 indicates default, 1 is explicitly single, >1 for multi-line spacing
  
  this.defaultLang = 1033;
  this.rtlfont = 0;
  this.ltrfont = 0;
  this.rtllang = 1033;
  this.ltrlang = 1033;
  this.rtlfontsize = 24;
  this.ltrfontsize = 24;
  this.colorrtl = 0;
  this.colorltr = 0;
  
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
			if ((type[i] == '-') && (i<type.length-1)) {
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
			this.doc.errors.push("Unknown RTF command "+cmd.type+", tried "+method);
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
		var hexstr = '';
		for (var i=0; i<this.hexStore.length; i++) {
			hexstr += this.hexStore[i].value;
		}
		if (this.group.charset == 'DBCS') {
			var err = "Attempt to decode "+hexstr+" as a double-byte character; will attempt to decode using CP1252";
			console.log(err);
			this.doc.errors.push(err);
			//Try CP1252 multi-byte encoding
			this.group.charset = 'CP1252';
		}
		this.group.addContent(new RTFSpan({
			value: iconv.decode(
				Buffer.from(hexstr, 'hex'), this.group.get('charset'))
		}));
		this.hexStore.splice(0);
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
	} else if (endingGroup !== undefined && endingGroup != null && endingGroup !== this.doc && !endingGroup.get('ignorable')) {
		for (var i=0; i<endingGroup.content.length; i++) {
			var item = endingGroup.content[i];
			this.addContent(doc, item);
		}
		process.emit('debug', 'GROUP END', endingGroup.type, endingGroup.get('ignorable'))
	}
  }
  this.ctrl$sect = function() {
	  //New section
	  this.flushHexStore();
	  if (this.group) this.groupStack.push(this.group);
	  this.group = new RTFGroup(this.group || this.doc);
  }
  this.ctrl$sectd = function() {
	  //Use "plain" properties as section defaults
	  this.group.style.fontSize = this.doc.getStyle('fontSize')
	  //When explicitly setting to plain, set all styles to false for group and span
      this.group.style.bold = false;
      this.group.style.italic = false;
      this.group.style.underline = false;
	  this.group.style.language = this.defaultLang;
      this.resetSpanStyle();
  }
  this.addContent = function(destination, content) {
		//Don't add any content while in a TxField
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
      //this.group = this.doc
	  this.group = new RTFGroup(this.doc);
    }
    //If there isn't already a style specified, use the current group style to start
    if (typeof cmd.style === 'undefined') {
        cmd.style = this.group.style;
    }
    //Update any styling specified for the current span
	if (this.spanStyle.caps) {
		cmd.value = cmd.value.toUpperCase();
	}
    var newSpan = new RTFSpan(cmd);
    newSpan.style.foreground = this.spanStyle.foreground;
    newSpan.style.bold = this.spanStyle.bold;
    newSpan.style.italic = this.spanStyle.italic;
	newSpan.style.direction = this.spanStyle.dir;
    //cmd.style.bold = this.spanStyle.bold;
    //cmd.style.italic = this.spanStyle.italic;
    //this.group.addContent(new RTFSpan(cmd));
    this.group.addContent(newSpan);
  }
  this.cmd$controlWord = function(cmd) {
    this.flushHexStore()
    if (typeof this.group !== 'undefined' && this.group !== null) {
      if (!this.group.type) this.group.type = cmd.value;
      	const method = this.deriveCmd(cmd.value, 'ctrl');
      	if (this[method]) {
        	this[method](cmd.param)
      	} else {
        	if (!this.group.get('ignorable')) {
				//process.emit('debug', method, cmd.param) 
				console.log("Unknown RTF control word "+cmd.value+", tried "+method);
				this.doc.errors.push("Unknown RTF control word "+cmd.value+", tried "+method);
			}
      	}
      }
  }
  this.cmd$hexchar = function(cmd) {
    this.hexStore.push(cmd)
  }
  this.cmd$error = function(cmd) {
    console.log('Error: ' + cmd.value + (cmd.row && cmd.col ? ' at line ' + cmd.row + ':' + cmd.col : '') + '.');
    //this.emit('error', new Error('Error: ' + cmd.value + (cmd.row && cmd.col ? ' at line ' + cmd.row + ':' + cmd.col : '') + '.'))
	this.doc.errors.push('Error: ' + cmd.value + (cmd.row && cmd.col ? ' at line ' + cmd.row + ':' + cmd.col : '') + '.');
  }

  this.ctrl$rtf = function() {
    //this.group = this.doc
  }

  // new line
  this.ctrl$line = function() {
    this.group.addContent(new RTFParagraph())
  }

  this.addAllStoredContent = function() {
		for (var i=0; i< this.groupStack.length; i++) {
			if ((this.groupStack[i] !== undefined) 
					&& (this.groupStack[i].content !== undefined) 
					&& (this.groupStack[i].content.length > 0)) {
				for (var j=0; j<this.groupStack[i].content.length; j++) {
					var item = this.groupStack[i].content[j];
					this.addContent(this.doc, item);
				}
			}
		}
		  
		this.flushHexStore();	
		if (this.group !== undefined && this.group.content !== undefined && this.group.content.length > 0) {
			this.addContent(this.doc, this.group);
		}
  }
  
  this.clearUnstoredContent = function() {
		this.flushHexStore();
		this.groupStack.length = 0;
		this.group = null;
  }

  this.ctrl$txfielddef = function() {
	  //Handle txfieldstart and texfieldend
	  this.inTxField = true;
  }

  this.ctrl$txfieldstart = function(cmd) {
	  //If we encounter a txfield, add all content that we've found thus far to the document,
	  //and flag to ignore anything inside the txfield
	  this.inTxField = true;
	  //Ignore the group that includes the txfieldstart
	  this.group.ignorable = true;
	  
  }

  this.ctrl$txfieldend = function(cmd) {
		//If we know we're in a txfield, resume collecting text
		if (this.inTxField) {
			//Ignore anything collected within txfield
			this.group.ignorable = true;
			
			//Clear txfield flag
			this.inTxField = false;
		}
		else {
			//If we encounter a txfieldend that was not preceded by a txfield start,
			//store everything so far and ignore everything after
			this.addAllStoredContent();
			this.abandonDocument = true;
		}
  }
  this.ctrl$viewkind = function(num) {
	  //Not supported; log request
	  var requestedView = '';
	  switch(num) {
			case 0: requestedView = 'None';
					break;
			case 1: requestedView = 'Page Layout';
					break;
			case 2: requestedView = 'Outline';
					break;
			case 3: requestedView = 'Master';
					break;
			case 4: requestedView = 'Normal';
					break;
			case 5: requestedView = 'Outline';
					break;
			default: requestedView = 'Unknown';
	  }
	  console.log("Requested document view "+requestedView);
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

  // general style
  this.resetSpanStyle = function() {
	this.spanStyle.bold = false;
    this.spanStyle.italic = false;
    this.spanStyle.foreground = {red: 0, blue: 0, green: 0};
	this.spanStyle.dir = 'ltr';
	this.spanStyle.caps = false;
  }
  this.nextParagraph = function(parentStyle) {
	this.group.addContent(new RTFParagraph(parentStyle));
	for (var i=1; i<this.spacing; i++) {
		//Add extra spacing if required
		this.group.addContent(new RTFParagraph(parentStyle));
	}
  }
  this.ctrl$par = function() {
    //Create new paragraph, starting from document styling
    this.nextParagraph(this.doc);
  }
  this.ctrl$n = function() {
    //For newlines, create new paragraph, starting from group styling
	if (!(this.group instanceof FontTable) && (!(this.group.parent instanceof FontTable))) {
		this.nextParagraph(this.group);
	}
  }
  this.ctrl$r = function() {
    //For carriage returns, create new paragraph, starting from group styling
	if (!(this.group instanceof FontTable) && (!(this.group.parent instanceof FontTable))) {
		this.nextParagraph(this.group);
	}
  }
  this.ctrl$pard = function() {
    this.group.resetStyle();
	//this.resetSpanStyle();
	this.spacing = 0;
  }
  this.ctrl$plain = function() {
    this.group.style.fontSize = this.doc.getStyle('fontSize')
    //When explicitly setting to plain, set all styles to false for group and span
    this.group.style.bold = false;
    this.group.style.italic = false;
    this.group.style.underline = false;
	this.group.style.language = this.defaultLang;
    this.resetSpanStyle();
	this.spacing = 0;
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
  this.ctrl$upr = function() {
	  this.group.ignorable = true;
  }
  this.ctrl$ud = function() {
	  //Indicated block containing Unicode
	  //Turn off any ignoring of previous upr section
	  this.group.ignorable = false;
	  if (this.group.parent !== undefined) {
		this.group.parent.ignorable = false;
	  }
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
  this.ctrl$strikedl = function(set) {
	 //Treat double strikethrough as single strikethrough
     this.group.style.strikethrough = set !== 0
  }
  this.ctrl$ul = function(set) {
    this.group.style.underline = set !== 0
  }
  this.ctrl$ulnone = function(set) {
    this.group.style.underline = false
  }
  this.ctrl$uld = function(set) {
	//Treat various underline styles as standard underline
    this.group.style.underline = set !== 0
  }
  this.ctrl$uldash = function(set) {
	//Treat various underline styles as standard underline
    this.group.style.underline = set !== 0
  }
  this.ctrl$uldashd = function(set) {
	//Treat various underline styles as standard underline
    this.group.style.underline = set !== 0
  }
  this.ctrl$uldashdd = function(set) {
	//Treat various underline styles as standard underline
    this.group.style.underline = set !== 0
  }
  this.ctrl$uldb = function(set) {
	//Treat various underline styles as standard underline
    this.group.style.underline = set !== 0
  }
  this.ctrl$ulth = function(set) {
	//Treat various underline styles as standard underline
    this.group.style.underline = set !== 0
  }
  this.ctrl$ulw = function(set) {
	//Treat various underline styles as standard underline
    this.group.style.underline = set !== 0
  }
  this.ctrl$ulwave = function(set) {
	//Treat various underline styles as standard underline
    this.group.style.underline = set !== 0
  }
  this.ctrl$caps = function() {
	  this.spanStyle.caps = true;
  }
  this.ctrl$scaps = function() {
	  //Treat small caps as regular all caps
	  this.spanStyle.caps = true;
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
  this.ctrl$ri = function(value) {
    this.group.style.rightindent = value;
  }
  
  this.ctrl$marglsxn = function(value) {
	  this.doc.style.indent = value;
  }
  this.ctrl$margrsxn = function(value) {
	  this.doc.style.rightindent = value;
  }
  this.ctrl$margtsxn = function(value) {
	  this.doc.style.padtop = value;
  }
  this.ctrl$margbsxn = function(value) {
	  this.doc.style.padbottom = value;
  }
  
  //Border options
  this.ctrl$pgbrdrhead = function(value) {
	  this.doc.style.borderheader = value * 100;
  }
  this.ctrl$pgbrdrfoot = function(value) {
	  this.doc.style.borderfooter = value * 100;
  }
  this.ctrl$pgbrdrt = function(value) {
	  this.doc.style.bordertop = value * 100;
  }
  this.ctrl$pgbrdrb = function(value) {
	  this.doc.style.borderbottom = value * 100;
  }
  this.ctrl$pgbrdrr = function(value) {
	  this.doc.style.borderright = value * 100;
  }
  this.ctrl$pgbrdrl = function(value) {
	  this.doc.style.borderleft = value * 100;
  }
  this.ctrl$pgbrdrart = function(value) {
	  //Not supported
  }
  this.ctrl$pgbrdropt = function(value) {
	  //Not supported
  }
  this.ctrl$pgbrdrsnap = function() {
	  //Not supported
  }
  
  //Special characters
  this.ctrl$tab = function() {
	  //Potentially vary tab width by adding multiple HTML tabs using this.tabWidth
      var spacer = { value: "&nbsp;", style: this.group.style };
      this.group.addContent(new RTFSpan(spacer));
  }
  this.ctrl$tx = function() {
	  //Setting tab stops not supported
  }
  this.ctrl$sa = function(num) {
	  this.group.style.padbottom = num;
  }
  this.ctrl$sl = function(num) {
	  this.spacing = 1;
  }
  this.ctrl$slmult = function(num) {
	  this.spacing = this.spacing*(num+1);
  }
  this.ctrl$deftab = function(num) {
	  this.defaultTabWidth = num;
  }
  this.ctrl$emdash = function() {
	  var emdash = { value: "&mdash; ", style: this.group.style };
      this.group.addContent(new RTFSpan(emdash));
  }
  this.ctrl$endash = function() {
	  var endash = { value: "&ndash; ", style: this.group.style };
      this.group.addContent(new RTFSpan(endash));
  }
  this.ctrl$emspace = function() {
	  var emspace = { value: "&emsp;", style: this.group.style };
      this.group.addContent(new RTFSpan(emspace));
  }
  this.ctrl$enspace = function() {
	  var enspace = { value: "&ensp;", style: this.group.style };
      this.group.addContent(new RTFSpan(enspace));
  }
  this.ctrl$bullet = function() {
	  var bullet = { value: "&bull; ", style: this.group.style };
      this.group.addContent(new RTFSpan(bullet));
  }
  this.ctrl$lquote = function() {
	  var quote = { value: "&lsquo;", style: this.group.style };
      this.group.addContent(new RTFSpan(quote));
  }
  this.ctrl$rquote = function() {
	  var quote = { value: "&rsquo;", style: this.group.style };
      this.group.addContent(new RTFSpan(quote));
  }
  this.ctrl$ldblquote = function() {
	  var quote = { value: "&ldquo;", style: this.group.style };
      this.group.addContent(new RTFSpan(quote));
  }
  this.ctrl$rdblquote = function() {
	  var quote = { value: "&rdquo;", style: this.group.style };
      this.group.addContent(new RTFSpan(quote));
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
		this.doc.errors.push('Error: ' + cmd.value + (cmd.row && cmd.col ? ' at line ' + cmd.row + ':' + cmd.col : '') + '.');
    } else {
      	this.group.charset = 'CP' + codepage
    }
  }
  this.ctrl$uc = function(n) {
	  //This works for n=1 (1 byte Unicode characters), different charsets may be needed for n>1
	  this.group.charset = 'CP1252';
	  if (n != 1) {
			var err = "Attempting to use "+n+" byte unicode; will attempt to decode using CP1252";
			console.log(err);
			this.doc.errors.push(err);
	  }
  }
  this.ctrl$hich = function() {
	  this.group.charset = 'CP1252';
  }
  this.ctrl$loch = function() {
	  if (this.savedCharset == '') {
		this.savedCharset = this.group.charset;
	  }
	  else {
		this.group.charset = this.savedCharset;
	  }
  }
  this.ctrl$dbch = function() {
	  this.group.charset = 'DBCS';
  }
  
  this.ctrl$itap = function(value) {
	  //Don't support paragraph nesting levels
  }
  
  //Explicitly ignore page-level controls mostly relevant to print settings, keeping text together on the printed page, etx
  this.ctrl$paperw = function() {
  }
  this.ctrl$paperh = function() {
  }
  this.ctrl$psz = function() {
  }
  this.ctrl$facingp = function() {
  }
  this.ctrl$margmirror = function() {
  }
  this.ctrl$landscape = function() {
  }
  this.ctrl$pgnstart = function() {
  }
  this.ctrl$widowctrl = function() {
  }
  this.ctrl$pgwsxn = function(num) {
  }
  this.ctrl$pghsxn = function(num) {
  }
  this.ctrl$hyphpar = function() {
  }
  this.ctrl$keep = function() {
  }
  this.ctrl$pagebb = function() {
  }
  this.ctrl$nowidctlpar = function() {
  }
  this.ctrl$widctlpar = function() {
  }
  this.ctrl$noproof = function() {
	  //Explicitly ignore tags indicating to ignore automatic proofreading
  }
  
  //Explicitly disregard form formatting
  this.ctrl$formprot = function() {
  }
  this.ctrl$allprot = function() {
  }
  this.ctrl$formshade = function() {
  }
  this.ctrl$formdisp = function() {
  }
  this.ctrl$printdata = function() {
  }
  
  this.ctrl$headery = function(num) {
	  this.doc.style.headerSpacing = num;
  }
  this.ctrl$footery = function(num) {
	  this.doc.style.footerSpacing = num;
  }
  // text direction
  this.ctrl$rtlch = function() {
    this.group.style.dir = 'rtl';
	this.group.style.font = this.fontrtl;
	this.group.style.fontSize = this.fontsizertl;
	this.spanStyle.foreground = this.colorrtl;
  }
  this.ctrl$ltrch = function() {
    this.group.style.dir = 'ltr';
	this.group.style.font = this.fontltr;
	this.group.style.fontSize = this.fontsizeltr;
	this.spanStyle.foreground = this.colorltr;
  }
  this.ctrl$rtldoc = function() {
    this.doc.style.dir = 'rtl';
	this.group.style.font = this.fontrtl;
	this.group.style.fontSize = this.fontsizertl;
	this.spanStyle.foreground = this.colorrtl;
  }
  this.ctrl$ltrdoc = function() {
    this.doc.style.dir = 'ltr';
	this.group.style.font = this.fontltr;
	this.group.style.fontSize = this.fontsizeltr;
	this.spanStyle.foreground = this.colorltr;
  }
  this.ctrl$rtlpar = function() {
    this.spanStyle.dir = 'rtl';
	this.group.style.font = this.fontrtl;
	this.group.style.fontSize = this.fontsizertl;
	this.spanStyle.foreground = this.colorrtl;
  }
  this.ctrl$ltrpar = function() {
    this.spanStyle.dir = 'ltr';
	this.group.style.font = this.fontltr;
	this.group.style.fontSize = this.fontsizeltr;
	this.spanStyle.foreground = this.colorltr;
  }
  this.ctrl$rtlmark = function() {
    this.spanStyle.dir = 'rtl'
	this.group.style.font = this.fontrtl;
	this.group.style.fontSize = this.fontsizertl;
	this.spanStyle.foreground = this.colorrtl;
  }
  this.ctrl$ltrmark = function() {
    this.spanStyle.dir = 'ltr';
	this.group.style.font = this.fontltr;
	this.group.style.fontSize = this.fontsizeltr;
	this.spanStyle.foreground = this.colorltr;
  }

  this.ctrl$af = function(value) {
	  if (this.group.style.dir == 'rtl') {
		  this.fontrtl = value;
	  }
	  else {
		  this.fontltr = value;
	  }
  }
  
  this.ctrl$alang = function(value) {
	  if (this.group.style.dir == 'rtl') {
		  this.langrtl = value;
	  }
	  else {
		  this.langltr = value;
	  }
  }
  
  this.ctrl$afs = function(value) {
	  if (this.group.style.dir == 'rtl') {
		  this.fontsizertl = value;
	  }
	  else {
		  this.fontsizeltr = value;
	  }
  }
  
  this.ctrl$acf = function(value) {
	  if (this.group.style.dir == 'rtl') {
		  this.colorrtl = value;
	  }
	  else {
		  this.colorltr = value;
	  }
  }
  
  //Unsupported features for bidrectional text
  this.ctrl$rin = function(value) {
	  //Don't support separate indenting for right-to-left paragraphs
  }
  this.ctrl$ab = function() {
  }
  this.ctrl$acaps = function() {
  }
  this.ctrl$adn = function(value) {
  }
  this.ctrl$afs = function() {
  }
  this.ctrl$ai = function() {
  }
  this.ctrl$aoutl = function() {
  }
  this.ctrl$ascaps = function() {
  }
  this.ctrl$ashad = function() {
  }
  this.ctrl$astrike = function() {
  }
  this.ctrl$aul = function() {
  }
  this.ctrl$auld = function() {
  }
  this.ctrl$auldb = function() {
  }
  this.ctrl$aulnone = function() {
  }
  this.ctrl$aulw = function() {
  }
  this.ctrl$aup = function(value) {
  }
  this.ctrl$adeff = function(value) {
  }
  this.ctrl$adeflang = function(value) {
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
  this.ctrl$deff = function(num) {
	  this.group.style.font = num;
  }
  this.ctrl$fnil = function() {
    if (this.group instanceof FontTable || this.group.parent instanceof FontTable) {
      this.group.get('currentFont').family = 'nil'
    }
  }
  this.ctrl$ftnil = function() {
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
		this.doc.errors.push('Unsupported charset code #' + code);
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
    //this.group.style.foreground = value
    this.spanStyle.foreground = value;
  }
  this.ctrl$cb = function(value) {
    this.group.style.background = value
  }
  this.ctrl$fs = function(value) {
    this.group.style.fontSize = value
  }
  
  this.ctrl$chshdng = function(value) {
	  //Don't support character shading
  }
  this.ctrl$chshpat = function(value) {
	  //Don't support background shading
  }
  this.ctrl$chcfpat = function(value) {
	  //Don't support background shading
  }
  this.ctrl$chcbpat = function(value) {
	  //Don't support color fill
  }
  
  this.ctrl$lang = function(value) {
	  this.group.style.language = value;
  }
  this.ctrl$deflang = function(value) {
	  this.defaultLang = value;
  }
  this.ctrl$deflangfe = function(value) {
	  //Don't support separate default and default Asian languages
  }
  this.ctrl$langfe = function(value) {
	  //Don't support separate default and default Asian languages
  }
  this.ctrl$langnp = function(value) {
	  //Don't support multiple languages
  }
  this.ctrl$langfenp = function(value) {
	  //Don't support multiple languages
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
  this.ctrl$s = function(value) {
		var err = "Attempt to reference stylesheet "+value+"; stylesheets not supported";
		console.log(err);
		this.doc.errors.push(err);
  }
  this.ctrl$listoverridetable = function() {
    this.group.ignorable = true
  }
  this.ctrl$ls = function(value) {
		var err = "Attempt to reference list override entry "+value+"; list override tables not supported";
		console.log(err);
		this.doc.errors.push(err);
  }
  this.ctrl$ilvl = function(value) {
		//Will need to be supported if list override formatting is supported in the future
		//For now, ignore; sould be preceded by an ls tag which will be logged
  }
  this.ctrl$listtext = function(value) {
		//listtext blocks will contain formatting that takes the place of \\ls blocks.
		//In the future if list overridge tables are supported, listtext blocks should be ignored.
		//For now, consume the listtext tag itself and simply follow any subsequent formatting tags
  }
  this.ctrl$info = function(value) {
    this.group.ignorable = true
  }
  this.ctrl$mmathPr = function(value) {
    this.group.ignorable = true
  }
  this.ctrl$filetbl = function(value) {
    this.group.ignorable = true
  }
  this.ctrl$pn = function(value) {
		//Ignore commands related to autoformatting of bullets and numbered lists
		this.group.ignorable = true
  }
  this.ctrl$pntext = function(value) {
		//pntext blocks will contain formatting that takes the place of \\pn blocks.
		//In the future if pn is supported, pntext blocks should be ignored.
		//For now, consume the pntext tag itself and simply follow any subsequent formatting tags
  }
  
  //Unsupported, ignore tag
  this.ctrl$lytexcttp = function() {
  }
  this.ctrl$viewbksp = function() {
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
