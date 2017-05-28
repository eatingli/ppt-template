'use strict';

Object.defineProperty(exports, "__esModule", {
	value: true
});

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _entities = require('entities');

var _entities2 = _interopRequireDefault(_entities);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var Slide = function () {
	function Slide(rel, content) {
		_classCallCheck(this, Slide);

		// ppt/slides/_rels/slideI.xml.rels
		this.rel = rel;

		// ppt/slides/slideI.xml
		this.content = content;
	}

	/**
  * 
  */


	_createClass(Slide, [{
		key: 'clone',
		value: function clone() {
			return new Slide(this.rel, this.content);
		}

		/**
   * 
   */

	}, {
		key: 'fill',
		value: function fill(word) {

			// 把 "&" 之類的符號轉換 &amp;  (XML Entities)
			var value = _entities2.default.encodeXML(word.value);
			var key = word.key;

			// offset: 避免遞迴置換...
			var offset = 0;
			var temp = 0;

			// Replace All
			while ((temp = this.content.indexOf(key, offset)) > -1) {

				this.content = replace(this.content, offset, key, value);
				offset = temp + value.length;
			}

			// return this;
		}

		/**
   * 
   */

	}, {
		key: 'fillAll',
		value: function fillAll(words) {
			var _this = this;

			words.forEach(function (word) {
				_this.fill(word);
			});
		}
	}]);

	return Slide;
}();

exports.default = Slide;


function replace(str, offset, a, b) {
	var index = str.indexOf(a, offset);
	return index > -1 ? str.substring(0, index) + str.substring(index, str.length).replace(a, b) : str;
}

// var test = "AABBCCAABB";
// console.log(replace(test, 3, 'A', 'XX')); // ---> AABBCCXXABB
// console.log(replace(test, 3, 'D', 'XX')); // ---> AABBCCAABB