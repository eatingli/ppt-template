var fs = require('fs');
var fsp = require('fs-promise');
var JSZip = require('jszip');
var xml2js = require('xml2js');
var Slide = require('./Slide');

var Presentation = module.exports = function() {
    this.contents = {};
}

Presentation.prototype.load = async function(pptx) {

    //解壓縮
    var zip = await JSZip.loadAsync(pptx);

    for (var key in zip.files) {

        //圖像之外都當作文字解析
        var ext = key.substr(key.lastIndexOf('.'));
        var type = (ext == '.xml' || ext == '.rels') ? 'string' : 'array';

        //將各檔案轉成字串，紀錄(檔名 : 純文字)
        var content = await zip.files[key].async(type);
        //console.log(key, ' ', content);

        this.contents[key] = content;
    }
}

Presentation.prototype.loadFile = async function(pptxFilePath) {

    var pptxFile = await fsp.readFile(pptxFilePath);
    await this.load(pptxFile);

}

Presentation.prototype.streamAs = async function(stream) {

    var newZip = JSZip();

    for (var key in this.contents) {
        if (this.contents[key]) newZip.file(key, this.contents[key]);
    }

    await generateNodeStreamAsync(stream, newZip);
};

Presentation.prototype.saveAs = async function(outputFilePath) {

    await this.streamAs(fs.createWriteStream(outputFilePath));
};

Presentation.prototype.getSlideCount = function() {
    return Object.keys(this.contents).filter(function(key) {
        return key.substr(0, 16) === "ppt/slides/slide"
    }).length;
}

Presentation.prototype.getSlide = function(index) {
    if (index < 1 || index > this.getSlideCount()) return null;
    var rel = this.contents['ppt/slides/_rels/slide' + index + '.xml.rels'];
    var content = this.contents['ppt/slides/slide' + index + '.xml'];
    return new Slide(rel, content);
}

Presentation.prototype.clone = function() {
    var newPresentation = new Presentation();
    newPresentation.contents = JSON.parse(JSON.stringify(this.contents))
    return newPresentation;
}

Presentation.prototype.generate = async function(slides) {

    var newPresentation = this.clone();
    var newContents = newPresentation.contents;
    var oldCount = newPresentation.getSlideCount();
    var builder = new xml2js.Builder();

    //清掉舊的 ppt/slides/slideX.xml & ppt/slides/_rels/slideX.xml.rels
    for (var i = 0; i < oldCount; i++) {
        delete newContents['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'];
        delete newContents['ppt/slides/slide' + (i + 1) + '.xml'];
    }

    //加入新的 ppt/slides/slideX.xml & ppt/slides/_rels/slideX.xml.rels
    for (var i = 0; i < slides.length; i++) {
        var slide = slides[i];
        newContents['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'] = slide.rel;
        newContents['ppt/slides/slide' + (i + 1) + '.xml'] = slide.content;
    }

    //修改 [Content_Types].xml
    var xmlJs = await xml2jsAsync(newPresentation.contents['[Content_Types].xml']);

    //清掉舊的
    xmlJs['Types']['Override'].forEach(function(override, index) {
        if (override['$'].PartName.substr(0, 17) == '/ppt/slides/slide') delete xmlJs['Types']['Override'][index];
    });

    //加入新的
    for (var i = 0; i < slides.length; i++) {
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
    var xmlJs1 = await xml2jsAsync(newPresentation.contents['ppt/_rels/presentation.xml.rels']);
    var xmlJs2 = await xml2jsAsync(newPresentation.contents['ppt/presentation.xml']);

    //刪除舊的
    xmlJs1['Relationships']['Relationship'].forEach(function(relationship, index) {
        if (relationship['$'].Target.substr(0, 12) == 'slides/slide') delete xmlJs1['Relationships']['Relationship'][index];
    });

    /*
    //整理rId
    xmlJs1['Relationships']['Relationship'].forEach(function(relationship, index) {
        if (relationship) relationship['$'].Id = 'rId' + (index + 1);
    });
    */

    //加入新的
    var rId = xmlJs1['Relationships']['Relationship'].length;
    for (var i = 1; i <= slides.length; i++) {
        xmlJs1['Relationships']['Relationship'].push({
            '$': {
                'Id': 'rId' + (rId + i),
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                'Target': 'slides/slide' + i + '.xml'
            }
        });
    }

    //計算 maxId
    var maxId = 0;
    xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach(function(ob) {
        if (+ob['$'].id > maxId) maxId = +ob['$'].id;
    });

    //刪除舊的
    xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach(function(ob, index) {
        delete xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'][index];
    });

    //加入新的
    for (var i = 1; i <= slides.length; i++) {
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
    var xmlJs = await xml2jsAsync(newPresentation.contents['docProps/app.xml']);
    xmlJs["Properties"]["Slides"][0] = slides.length;
    newContents['docProps/app.xml'] = builder.buildObject(xmlJs);

    return new Promise((resolve, reject) => {
        resolve(newPresentation);
    });
}

//包裝成Promise版本
async function xml2jsAsync(xml) {
    return new Promise((resolve, reject) => {
        xml2js.parseString(xml, (err, xmlJs) => {
            if (err) throw reject(err);
            else resolve(xmlJs);
        });
    });
}

async function generateNodeStreamAsync(stream, zip) {
    return new Promise((resolve, reject) => {
        zip
            .generateNodeStream({
                type: 'nodebuffer',
                streamFiles: true
            })
            .pipe(stream)
            .on('finish', () => {
                resolve();
            });
    });

}


