'use strict'

class RTFSpan {
  constructor (opts) {
    if (!opts) opts = {}
    this.value = opts.value
    this.style = Object.assign({}, opts.style)
  }
}

module.exports = RTFSpan
