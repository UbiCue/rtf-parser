'use strict'

function RTFGroup(parent) {
  this.parent = parent
  this.content = []
  this.fonts = []
  if (typeof parent !== 'undefined' && typeof parent.colors !== 'undefined') {
      this.colors = parent.colors;
  }
  else {
      this.colors = [];
  }
  if (typeof parent != 'undefined' && typeof parent.style !== 'undefined') {
      this.style = parent.style;
  }
  else {
      this.style = {};
  }
  this.ignorable = null

  this.get = function(name) {
    return this[name] != null ? this[name] : this.parent.get(name)
  };
  this.getFont = function(num) {
    return this.fonts[num] != null && this.fonts[num] !== undefined ? this.fonts[num] : this.parent.getFont(num)
  };
  this.getColor = function(num) {
    return this.colors[num] != null ? this.colors[num] : this.parent.getFont(num)
  };
  this.getStyle = function(name) {
    if (!name) return Object.assign({}, this.parent.getStyle(), this.style)
    return this.style[name] != null ? this.style[name] : this.parent.getStyle(name)
  };
  this.resetStyle = function() {
    this.style = {}
  };
  this.addContent = function(node) {
    if (typeof node.style == 'undefined') {
        node.style = Object.assign({}, this.getStyle());
    }
    if (typeof node.style.font != 'undefined') {
        //If we're not already being passed a Font object, retrieve a Font by numerical value
        if (typeof node.style.font.name == 'undefined') {
            node.style.font = this.getFont(node.style.font);
        };
    }
    if (typeof node.style.foreground != 'undefined' && node.style.foreground != null) {
        if (typeof node.style.foreground.red === 'undefined') {
            node.style.foreground = this.getColor(node.style.foreground);
        }
    }
    if (typeof node.style.background != 'undefined' && node.style.background != null) {
        if (typeof node.style.background.red === 'undefined') {
            node.style.background = this.getColor(node.style.background);
        }
    }
    this.content.push(node)
  };
};

module.exports = RTFGroup
