'use strict';

//包裝成Promise版本
var xml2jsAsync = function () {
    var _ref6 = _asyncToGenerator(regeneratorRuntime.mark(function _callee6(xml) {
        return regeneratorRuntime.wrap(function _callee6$(_context6) {
            while (1) {
                switch (_context6.prev = _context6.next) {
                    case 0:
                        return _context6.abrupt('return', new Promise(function (resolve, reject) {
                            xml2js.parseString(xml, function (err, xmlJs) {
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

/*
Presentation.prototype.getXml2Js = function(key, callback) {
    xml2js.parseString(this.contents[key], function(err, xmlJs) {
        if (err) throw err;
        callback(xmlJs);
    });
}

Presentation.prototype.getJs2Xml = getJs2Xml;

function getJs2Xml(js) {
    return new xml2js.Builder().buildObject(js);
}
*/


function _asyncToGenerator(fn) { return function () { var gen = fn.apply(this, arguments); return new Promise(function (resolve, reject) { function step(key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { return Promise.resolve(value).then(function (value) { step("next", value); }, function (err) { step("throw", err); }); } } return step("next"); }); }; }

var fs = require('fs');
var fsp = require('fs-promise');
var JSZip = require('jszip');
var xml2js = require('xml2js');
var Slide = require('./Slide');

var Presentation = module.exports = function () {
    this.contents = {};
};

Presentation.prototype.load = function () {
    var _ref = _asyncToGenerator(regeneratorRuntime.mark(function _callee(pptx) {
        var zip, key, ext, type, content;
        return regeneratorRuntime.wrap(function _callee$(_context) {
            while (1) {
                switch (_context.prev = _context.next) {
                    case 0:
                        _context.next = 2;
                        return JSZip.loadAsync(pptx);

                    case 2:
                        zip = _context.sent;
                        _context.t0 = regeneratorRuntime.keys(zip.files);

                    case 4:
                        if ((_context.t1 = _context.t0()).done) {
                            _context.next = 14;
                            break;
                        }

                        key = _context.t1.value;


                        //圖像之外都當作文字解析
                        ext = key.substr(key.lastIndexOf('.'));
                        type = ext == '.xml' || ext == '.rels' ? 'string' : 'array';

                        //將各檔案轉成字串，紀錄(檔名 : 純文字)

                        _context.next = 10;
                        return zip.files[key].async(type);

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

    return function (_x) {
        return _ref.apply(this, arguments);
    };
}();

Presentation.prototype.loadFile = function () {
    var _ref2 = _asyncToGenerator(regeneratorRuntime.mark(function _callee2(pptxFilePath) {
        var pptxFile;
        return regeneratorRuntime.wrap(function _callee2$(_context2) {
            while (1) {
                switch (_context2.prev = _context2.next) {
                    case 0:
                        _context2.next = 2;
                        return fsp.readFile(pptxFilePath);

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

    return function (_x2) {
        return _ref2.apply(this, arguments);
    };
}();

Presentation.prototype.streamAs = function () {
    var _ref3 = _asyncToGenerator(regeneratorRuntime.mark(function _callee3(stream) {
        var newZip, key;
        return regeneratorRuntime.wrap(function _callee3$(_context3) {
            while (1) {
                switch (_context3.prev = _context3.next) {
                    case 0:
                        newZip = JSZip();


                        for (key in this.contents) {
                            if (this.contents[key]) newZip.file(key, this.contents[key]);
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

    return function (_x3) {
        return _ref3.apply(this, arguments);
    };
}();

Presentation.prototype.saveAs = function () {
    var _ref4 = _asyncToGenerator(regeneratorRuntime.mark(function _callee4(outputFilePath) {
        return regeneratorRuntime.wrap(function _callee4$(_context4) {
            while (1) {
                switch (_context4.prev = _context4.next) {
                    case 0:
                        _context4.next = 2;
                        return this.streamAs(fs.createWriteStream(outputFilePath));

                    case 2:
                    case 'end':
                        return _context4.stop();
                }
            }
        }, _callee4, this);
    }));

    return function (_x4) {
        return _ref4.apply(this, arguments);
    };
}();

Presentation.prototype.getSlideCount = function () {
    return Object.keys(this.contents).filter(function (key) {
        return key.substr(0, 16) === "ppt/slides/slide";
    }).length;
};

Presentation.prototype.getSlide = function (index) {
    if (index < 1 || index > this.getSlideCount()) return null;
    var rel = this.contents['ppt/slides/_rels/slide' + index + '.xml.rels'];
    var content = this.contents['ppt/slides/slide' + index + '.xml'];
    return new Slide(rel, content);
};

Presentation.prototype.clone = function () {
    var newPresentation = new Presentation();
    newPresentation.contents = JSON.parse(JSON.stringify(this.contents));
    return newPresentation;
};

Presentation.prototype.generate = function () {
    var _ref5 = _asyncToGenerator(regeneratorRuntime.mark(function _callee5(slides) {
        var newPresentation, newContents, oldCount, builder, i, slide, xmlJs, xmlJs1, xmlJs2, rId, maxId;
        return regeneratorRuntime.wrap(function _callee5$(_context5) {
            while (1) {
                switch (_context5.prev = _context5.next) {
                    case 0:
                        newPresentation = this.clone();
                        newContents = newPresentation.contents;
                        oldCount = newPresentation.getSlideCount();
                        builder = new xml2js.Builder();

                        //清掉舊的 ppt/slides/slideX.xml & ppt/slides/_rels/slideX.xml.rels

                        for (i = 0; i < oldCount; i++) {
                            delete newContents['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'];
                            delete newContents['ppt/slides/slide' + (i + 1) + '.xml'];
                        }

                        //加入新的 ppt/slides/slideX.xml & ppt/slides/_rels/slideX.xml.rels
                        for (i = 0; i < slides.length; i++) {
                            slide = slides[i];

                            newContents['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'] = slide.rel;
                            newContents['ppt/slides/slide' + (i + 1) + '.xml'] = slide.content;
                        }

                        //修改 [Content_Types].xml
                        _context5.next = 8;
                        return xml2jsAsync(newPresentation.contents['[Content_Types].xml']);

                    case 8:
                        xmlJs = _context5.sent;


                        //清掉舊的
                        xmlJs['Types']['Override'].forEach(function (override, index) {
                            if (override['$'].PartName.substr(0, 17) == '/ppt/slides/slide') delete xmlJs['Types']['Override'][index];
                        });

                        //加入新的
                        for (i = 0; i < slides.length; i++) {
                            xmlJs['Types']['Override'].push({
                                '$': {
                                    'PartName': '/ppt/slides/slide' + (i + 1) + '.xml',
                                    'ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
                                }
                            });
                        }

                        //更新
                        newContents['[Content_Types].xml'] = builder.buildObject(xmlJs);

                        //修改 ppt/_rels/presentation.xml.rels & ppt/presentation.xml
                        _context5.next = 14;
                        return xml2jsAsync(newPresentation.contents['ppt/_rels/presentation.xml.rels']);

                    case 14:
                        xmlJs1 = _context5.sent;
                        _context5.next = 17;
                        return xml2jsAsync(newPresentation.contents['ppt/presentation.xml']);

                    case 17:
                        xmlJs2 = _context5.sent;


                        //刪除舊的
                        xmlJs1['Relationships']['Relationship'].forEach(function (relationship, index) {
                            if (relationship['$'].Target.substr(0, 12) == 'slides/slide') delete xmlJs1['Relationships']['Relationship'][index];
                        });

                        /*
                        //整理rId
                        xmlJs1['Relationships']['Relationship'].forEach(function(relationship, index) {
                            if (relationship) relationship['$'].Id = 'rId' + (index + 1);
                        });
                        */

                        //加入新的
                        rId = xmlJs1['Relationships']['Relationship'].length;

                        for (i = 1; i <= slides.length; i++) {
                            xmlJs1['Relationships']['Relationship'].push({
                                '$': {
                                    'Id': 'rId' + (rId + i),
                                    'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                                    'Target': 'slides/slide' + i + '.xml'
                                }
                            });
                        }

                        //計算 maxId
                        maxId = 0;

                        xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach(function (ob) {
                            if (+ob['$'].id > maxId) maxId = +ob['$'].id;
                        });

                        //刪除舊的
                        xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach(function (ob, index) {
                            delete xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'][index];
                        });

                        //加入新的
                        for (i = 1; i <= slides.length; i++) {
                            xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'].push({
                                '$': {
                                    'id': '' + (+maxId + i),
                                    'r:id': 'rId' + (rId + i)
                                }
                            });
                        }

                        newContents['ppt/_rels/presentation.xml.rels'] = builder.buildObject(xmlJs1);
                        newContents['ppt/presentation.xml'] = builder.buildObject(xmlJs2);

                        //修改 docProps/app.xml (increment slidecount)
                        _context5.next = 29;
                        return xml2jsAsync(newPresentation.contents['docProps/app.xml']);

                    case 29:
                        xmlJs = _context5.sent;

                        xmlJs["Properties"]["Slides"][0] = slides.length;
                        newContents['docProps/app.xml'] = builder.buildObject(xmlJs);

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

    return function (_x5) {
        return _ref5.apply(this, arguments);
    };
}();