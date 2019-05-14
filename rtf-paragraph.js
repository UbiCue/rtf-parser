'use strict'

function RTFParagraph(opts) {
  if (!opts) opts = {}
  this.style = opts.style || {}
  this.content = []
}

module.exports = RTFParagraph
