'use strict'
const RTFGroup = require('./rtf-group.js')
const RTFParagraph = require('./rtf-paragraph.js')

function RTFDocument() {
  RTFGroup.call(this, {});
  this.charset = 'ASCII'
  this.ignorable = false
  this.marginLeft = 1800
  this.marginRight = 1800
  this.marginBottom = 1440
  this.marginTop = 1440
  this.style = {
    font: 0,
    fontSize: 24,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    foreground: null,
    background: null,
    firstLineIndent: 0,
    indent: 0,
    align: 'left',
    valign: 'normal'
  }
  this.get = function(name) {
    return this[name]
  }
  this.getFont = function(num) {
    return this.fonts[num]
  }
  this.getColor = function(num) {
    return this.colors[num]
  }
  this.getStyle = function(name) {
    if (!name) return this.style
    return this.style[name]
  }
  this.docAddContent = function(node) {
    if (node instanceof RTFParagraph) {
      while (this.content.length && !(this.content[this.content.length - 1] instanceof RTFParagraph)) {
        node.content.unshift(this.content.pop())
      }
      this.addContent(node)
      if (node.content.length) {
        const initialStyle = node.content[0].style
        var style = {}
        if (typeof node.style !== 'undefined' && node.style != null) {
            style = node.style;
        }
        else {
            style = {};
        }
        if (typeof style.font === 'undefined' || style.font == null) {
            style.font = this.getFont(initialStyle.font);
        }
        if (typeof style.foreground === 'undefined' || style.foreground == null) {
            style.foreground = this.getColor(initialStyle.foreground);
        }
        if (typeof style.background === 'undefined' || style.background == null) {
            style.background = this.getColor(initialStyle.background);
        }
        for (prop of Object.keys(initialStyle)) {
          if (initialStyle[prop] == null) continue
          let match = true
          for (span of node.content) {
            if (initialStyle[prop] !== span.style[prop]) {
              match = false
              break
            }
          }
          if (match) style[prop] = initialStyle[prop]
        }
        node.style = style
      }
    } else {
      this.addContent(node)
    }
  }
}

module.exports = RTFDocument
