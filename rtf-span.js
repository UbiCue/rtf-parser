'use strict'

function RTFSpan(opts) {
  if (!opts) opts = {}
  this.value = opts.value
  for (var attr in opts.style) {
    this.style[attr] = opts.style[attr];
  }
}

module.exports = RTFSpan
