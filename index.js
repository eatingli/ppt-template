require('babel-polyfill');

const package = {};
package.Presentation = require('./build/Presentation').default;
package.Slide = require('./build/Slide').default;

module.exports = package;