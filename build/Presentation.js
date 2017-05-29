'use strict';

Object.defineProperty(exports, "__esModule", {
    value: true
});

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

/**
 * Parse xml to object (Promise)
 */
var xml2jsAsync = function () {
    var _ref6 = _asyncToGenerator(regeneratorRuntime.mark(function _callee6(xml) {
        return regeneratorRuntime.wrap(function _callee6$(_context6) {
            while (1) {
                switch (_context6.prev = _context6.next) {
                    case 0:
                        return _context6.abrupt('return', new Promise(function (resolve, reject) {
                            _xml2js2.default.parseString(xml, function (err, xmlJs) {
                                if (err) throw reject(err);else resolve(xmlJs);
                            });
                        }));

                    case 1:
                    case 'end':
                        return _context6.stop();
                }
            }
        }, _callee6, this);
    }));

    return function xml2jsAsync(_x6) {
        return _ref6.apply(this, arguments);
    };
}();

/**
 * generateNodeStreamAsync (Promise)
 */


var generateNodeStreamAsync = function () {
    var _ref7 = _asyncToGenerator(regeneratorRuntime.mark(function _callee7(stream, zip) {
        return regeneratorRuntime.wrap(function _callee7$(_context7) {
            while (1) {
                switch (_context7.prev = _context7.next) {
                    case 0:
                        return _context7.abrupt('return', new Promise(function (resolve, reject) {
                            zip.generateNodeStream({
                                type: 'nodebuffer',
                                streamFiles: true
                            }).pipe(stream).on('finish', function () {
                                resolve();
                            });
                        }));

                    case 1:
                    case 'end':
                        return _context7.stop();
                }
            }
        }, _callee7, this);
    }));

    return function generateNodeStreamAsync(_x7, _x8) {
        return _ref7.apply(this, arguments);
    };
}();

var _fs = require('fs');

var _fs2 = _interopRequireDefault(_fs);

var _fsPromise = require('fs-promise');

var _fsPromise2 = _interopRequireDefault(_fsPromise);

var _jszip = require('jszip');

var _jszip2 = _interopRequireDefault(_jszip);

var _xml2js = require('xml2js');

var _xml2js2 = _interopRequireDefault(_xml2js);

var _Slide = require('./Slide.js');

var _Slide2 = _interopRequireDefault(_Slide);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _asyncToGenerator(fn) { return function () { var gen = fn.apply(this, arguments); return new Promise(function (resolve, reject) { function step(key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { return Promise.resolve(value).then(function (value) { step("next", value); }, function (err) { step("throw", err); }); } } return step("next"); }); }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var Presentation = function () {
    function Presentation() {
        _classCallCheck(this, Presentation);

        this.contents = {};
    }

    /**
     * Load the .pptx stream
     */


    _createClass(Presentation, [{
        key: 'load',
        value: function () {
            var _ref = _asyncToGenerator(regeneratorRuntime.mark(function _callee(stream) {
                var pptx, key, ext, type, content;
                return regeneratorRuntime.wrap(function _callee$(_context) {
                    while (1) {
                        switch (_context.prev = _context.next) {
                            case 0:
                                _context.next = 2;
                                return _jszip2.default.loadAsync(stream);

                            case 2:
                                pptx = _context.sent;
                                _context.t0 = regeneratorRuntime.keys(pptx.files);

                            case 4:
                                if ((_context.t1 = _context.t0()).done) {
                                    _context.next = 14;
                                    break;
                                }

                                key = _context.t1.value;


                                // 圖像之外都當作文字解析
                                ext = key.substr(key.lastIndexOf('.'));
                                type = ext == '.xml' || ext == '.rels' ? 'string' : 'array';

                                // 將各檔案轉成字串

                                _context.next = 10;
                                return pptx.files[key].async(type);

                            case 10:
                                content = _context.sent;

                                //console.log(key, ' ', content);

                                this.contents[key] = content;
                                _context.next = 4;
                                break;

                            case 14:
                            case 'end':
                                return _context.stop();
                        }
                    }
                }, _callee, this);
            }));

            function load(_x) {
                return _ref.apply(this, arguments);
            }

            return load;
        }()

        /**
         * Load the .pptx file
         */

    }, {
        key: 'loadFile',
        value: function () {
            var _ref2 = _asyncToGenerator(regeneratorRuntime.mark(function _callee2(file) {
                var pptxFile;
                return regeneratorRuntime.wrap(function _callee2$(_context2) {
                    while (1) {
                        switch (_context2.prev = _context2.next) {
                            case 0:
                                _context2.next = 2;
                                return _fsPromise2.default.readFile(file);

                            case 2:
                                pptxFile = _context2.sent;
                                _context2.next = 5;
                                return this.load(pptxFile);

                            case 5:
                            case 'end':
                                return _context2.stop();
                        }
                    }
                }, _callee2, this);
            }));

            function loadFile(_x2) {
                return _ref2.apply(this, arguments);
            }

            return loadFile;
        }()

        /**
         * 
         */

    }, {
        key: 'streamAs',
        value: function () {
            var _ref3 = _asyncToGenerator(regeneratorRuntime.mark(function _callee3(stream) {
                var newZip, key;
                return regeneratorRuntime.wrap(function _callee3$(_context3) {
                    while (1) {
                        switch (_context3.prev = _context3.next) {
                            case 0:
                                newZip = (0, _jszip2.default)();


                                for (key in this.contents) {
                                    if (this.contents[key]) newZip.file(key, this.contents[key]);else console.error('No content', key);
                                }

                                _context3.next = 4;
                                return generateNodeStreamAsync(stream, newZip);

                            case 4:
                            case 'end':
                                return _context3.stop();
                        }
                    }
                }, _callee3, this);
            }));

            function streamAs(_x3) {
                return _ref3.apply(this, arguments);
            }

            return streamAs;
        }()

        /**
         * 
         */

    }, {
        key: 'saveAs',
        value: function () {
            var _ref4 = _asyncToGenerator(regeneratorRuntime.mark(function _callee4(file) {
                return regeneratorRuntime.wrap(function _callee4$(_context4) {
                    while (1) {
                        switch (_context4.prev = _context4.next) {
                            case 0:
                                _context4.next = 2;
                                return this.streamAs(_fs2.default.createWriteStream(file));

                            case 2:
                            case 'end':
                                return _context4.stop();
                        }
                    }
                }, _callee4, this);
            }));

            function saveAs(_x4) {
                return _ref4.apply(this, arguments);
            }

            return saveAs;
        }()

        /**
         * Get slide amount.
         */

    }, {
        key: 'getSlideCount',
        value: function getSlideCount() {
            return Object.keys(this.contents).filter(function (key) {
                return key.substr(0, 16) === 'ppt/slides/slide';
            }).length;
        }

        /**
         * Get silde by index, index is from 1 to length.
         */

    }, {
        key: 'getSlide',
        value: function getSlide(index) {

            if (index < 1 || index > this.getSlideCount()) return null;

            var rel = this.contents['ppt/slides/_rels/slide' + index + '.xml.rels'];
            var content = this.contents['ppt/slides/slide' + index + '.xml'];
            return new _Slide2.default(rel, content);
        }

        /**
         * Clone presention
         */

    }, {
        key: 'clone',
        value: function clone() {
            var newPresentation = new Presentation();
            newPresentation.contents = JSON.parse(JSON.stringify(this.contents));
            return newPresentation;
        }

        /**
         * 
         */

    }, {
        key: 'generate',
        value: function () {
            var _ref5 = _asyncToGenerator(regeneratorRuntime.mark(function _callee5(slides) {
                var newPresentation, newContents, slideCount, builder, i, _i, slide, temp, _i2, rId, _i3, maxId, _i4;

                return regeneratorRuntime.wrap(function _callee5$(_context5) {
                    while (1) {
                        switch (_context5.prev = _context5.next) {
                            case 0:
                                newPresentation = this.clone();
                                newContents = newPresentation.contents;
                                slideCount = newPresentation.getSlideCount();
                                builder = new _xml2js2.default.Builder();

                                // Clear "ppt/slides/slideX.xml" & "ppt/slides/_rels/slideX.xml.rels"

                                for (i = 0; i < slideCount; i++) {
                                    delete newContents['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'];
                                    delete newContents['ppt/slides/slide' + (i + 1) + '.xml'];
                                }

                                // 加入新的 ppt/slides/slideX.xml & ppt/slides/_rels/slideX.xml.rels
                                for (_i = 0; _i < slides.length; _i++) {
                                    slide = slides[_i];

                                    newContents['ppt/slides/_rels/slide' + (_i + 1) + '.xml.rels'] = slide.rel;
                                    newContents['ppt/slides/slide' + (_i + 1) + '.xml'] = slide.content;
                                }

                                //# Edit "[Content_Types].xml""
                                _context5.next = 8;
                                return xml2jsAsync(newContents['[Content_Types].xml']);

                            case 8:
                                temp = _context5.sent;


                                // Clear old
                                temp['Types']['Override'].forEach(function (override, index) {
                                    if (override['$'].PartName.substr(0, 17) == '/ppt/slides/slide') delete temp['Types']['Override'][index];
                                });

                                // 加入新的
                                for (_i2 = 0; _i2 < slides.length; _i2++) {
                                    temp['Types']['Override'].push({
                                        '$': {
                                            'PartName': '/ppt/slides/slide' + (_i2 + 1) + '.xml',
                                            'ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
                                        }
                                    });
                                }

                                //更新
                                newContents['[Content_Types].xml'] = builder.buildObject(temp);

                                //# 修改 ppt/_rels/presentation.xml.rels
                                _context5.next = 14;
                                return xml2jsAsync(newContents['ppt/_rels/presentation.xml.rels']);

                            case 14:
                                temp = _context5.sent;


                                //刪除舊的
                                temp['Relationships']['Relationship'].forEach(function (relationship, index) {
                                    if (relationship['$'].Target.substr(0, 12) == 'slides/slide') delete temp['Relationships']['Relationship'][index];
                                });

                                /*
                                //整理rId
                                temp['Relationships']['Relationship'].forEach(function(relationship, index) {
                                    if (relationship) relationship['$'].Id = 'rId' + (index + 1);
                                });
                                */

                                // 加入新的
                                rId = temp['Relationships']['Relationship'].length;

                                for (_i3 = 1; _i3 <= slides.length; _i3++) {
                                    temp['Relationships']['Relationship'].push({
                                        '$': {
                                            'Id': 'rId' + (rId + _i3),
                                            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                                            'Target': 'slides/slide' + _i3 + '.xml'
                                        }
                                    });
                                }

                                newContents['ppt/_rels/presentation.xml.rels'] = builder.buildObject(temp);

                                //# 修改 ppt/presentation.xml
                                _context5.next = 21;
                                return xml2jsAsync(newContents['ppt/presentation.xml']);

                            case 21:
                                temp = _context5.sent;


                                //計算 maxId
                                maxId = 0;

                                temp['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach(function (ob) {
                                    if (+ob['$'].id > maxId) maxId = +ob['$'].id;
                                });

                                // 刪除舊的
                                temp['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach(function (ob, index) {
                                    delete temp['p:presentation']['p:sldIdLst'][0]['p:sldId'][index];
                                });

                                // 加入新的
                                for (_i4 = 1; _i4 <= slides.length; _i4++) {
                                    temp['p:presentation']['p:sldIdLst'][0]['p:sldId'].push({
                                        '$': {
                                            'id': '' + (+maxId + _i4),
                                            'r:id': 'rId' + (rId + _i4)
                                        }
                                    });
                                }

                                newContents['ppt/presentation.xml'] = builder.buildObject(temp);

                                // 修改 docProps/app.xml (increment slidecount)
                                _context5.next = 29;
                                return xml2jsAsync(newContents['docProps/app.xml']);

                            case 29:
                                temp = _context5.sent;

                                temp["Properties"]["Slides"][0] = slides.length;
                                newContents['docProps/app.xml'] = builder.buildObject(temp);

                                return _context5.abrupt('return', new Promise(function (resolve, reject) {
                                    resolve(newPresentation);
                                }));

                            case 33:
                            case 'end':
                                return _context5.stop();
                        }
                    }
                }, _callee5, this);
            }));

            function generate(_x5) {
                return _ref5.apply(this, arguments);
            }

            return generate;
        }()
    }]);

    return Presentation;
}();

exports.default = Presentation;