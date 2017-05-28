
import fs from 'fs'
import fsp from 'fs-promise'
import JSZip from 'jszip'
import xml2js from 'xml2js'

import Slide from './Slide.js'

export default class Presentation {

    constructor() {
        this.contents = {};
    }

    /**
     * Load the .pptx stream
     */
    async load(stream) {

        // Zip file
        let zip = await JSZip.loadAsync(stream);

        for (let key in zip.files) {

            // 圖像之外都當作文字解析
            let ext = key.substr(key.lastIndexOf('.'));
            let type = (ext == '.xml' || ext == '.rels') ? 'string' : 'array';

            // 將各檔案轉成字串
            let content = await zip.files[key].async(type);
            //console.log(key, ' ', content);

            this.contents[key] = content;
        }
    }

    /**
     * Load the .pptx file
     */
    async loadFile(file) {
        let pptxFile = await fsp.readFile(file);
        await this.load(pptxFile);
    }

    /**
     * 
     */
    async streamAs(stream) {
        let newZip = JSZip();

        for (let key in this.contents) {
            if (this.contents[key]) newZip.file(key, this.contents[key]);
        }

        await generateNodeStreamAsync(stream, newZip);
    }

    /**
     * 
     */
    async saveAs(file) {
        await this.streamAs(fs.createWriteStream(file));
    }

    /**
     * Get slide amount.
     */
    getSlideCount() {
        return Object.keys(this.contents).filter((key) => {
            return key.substr(0, 16) === 'ppt/slides/slide'
        }).length;
    }

    /**
     * Get silde by index, index is from 1 to length.
     */
    getSlide(index) {
        if (index < 1 || index > this.getSlideCount()) return null;
        let rel = this.contents['ppt/slides/_rels/slide' + index + '.xml.rels'];
        let content = this.contents['ppt/slides/slide' + index + '.xml'];
        return new Slide(rel, content);
    }

    /**
     * 
     */
    clone() {
        let newPresentation = new Presentation();
        newPresentation.contents = JSON.parse(JSON.stringify(this.contents))
        return newPresentation;
    }

    /**
     * 
     */
    async generate(slides) {
        let newPresentation = this.clone();
        let newContents = newPresentation.contents;
        let oldCount = newPresentation.getSlideCount();
        let builder = new xml2js.Builder();

        // 清掉舊的 ppt/slides/slideX.xml & ppt/slides/_rels/slideX.xml.rels
        for (let i = 0; i < oldCount; i++) {
            delete newContents['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'];
            delete newContents['ppt/slides/slide' + (i + 1) + '.xml'];
        }

        // 加入新的 ppt/slides/slideX.xml & ppt/slides/_rels/slideX.xml.rels
        for (let i = 0; i < slides.length; i++) {
            let slide = slides[i];
            newContents['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'] = slide.rel;
            newContents['ppt/slides/slide' + (i + 1) + '.xml'] = slide.content;
        }

        //# 修改 [Content_Types].xml
        let temp = await xml2jsAsync(newPresentation.contents['[Content_Types].xml']);

        // 清掉舊的
        temp['Types']['Override'].forEach((override, index) => {
            if (override['$'].PartName.substr(0, 17) == '/ppt/slides/slide')
                delete temp['Types']['Override'][index];
        });

        // 加入新的
        for (let i = 0; i < slides.length; i++) {
            temp['Types']['Override'].push({
                '$': {
                    'PartName': '/ppt/slides/slide' + (i + 1) + '.xml',
                    'ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
                }
            });
        }

        //更新
        newContents['[Content_Types].xml'] = builder.buildObject(temp);

        //# 修改 ppt/_rels/presentation.xml.rels
        temp = await xml2jsAsync(newPresentation.contents['ppt/_rels/presentation.xml.rels']);

        //刪除舊的
        temp['Relationships']['Relationship'].forEach((relationship, index) => {
            if (relationship['$'].Target.substr(0, 12) == 'slides/slide')
                delete temp['Relationships']['Relationship'][index];
        });

        /*
        //整理rId
        temp['Relationships']['Relationship'].forEach(function(relationship, index) {
            if (relationship) relationship['$'].Id = 'rId' + (index + 1);
        });
        */

        // 加入新的
        let rId = temp['Relationships']['Relationship'].length;
        for (let i = 1; i <= slides.length; i++) {
            temp['Relationships']['Relationship'].push({
                '$': {
                    'Id': 'rId' + (rId + i),
                    'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                    'Target': 'slides/slide' + i + '.xml'
                }
            });
        }

        newContents['ppt/_rels/presentation.xml.rels'] = builder.buildObject(temp);

        //# 修改 ppt/presentation.xml
        temp = await xml2jsAsync(newPresentation.contents['ppt/presentation.xml']);

        //計算 maxId
        let maxId = 0;
        temp['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach((ob) => {
            if (+ob['$'].id > maxId) maxId = +ob['$'].id;
        });

        // 刪除舊的
        temp['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach((ob, index) => {
            delete temp['p:presentation']['p:sldIdLst'][0]['p:sldId'][index];
        });

        // 加入新的
        for (let i = 1; i <= slides.length; i++) {
            temp['p:presentation']['p:sldIdLst'][0]['p:sldId'].push({
                '$': {
                    'id': '' + (+maxId + i),
                    'r:id': 'rId' + (rId + i)
                }
            });
        }

        newContents['ppt/presentation.xml'] = builder.buildObject(temp);

        // 修改 docProps/app.xml (increment slidecount)
        temp = await xml2jsAsync(newPresentation.contents['docProps/app.xml']);
        temp["Properties"]["Slides"][0] = slides.length;
        newContents['docProps/app.xml'] = builder.buildObject(temp);

        return new Promise((resolve, reject) => {
            resolve(newPresentation);
        });
    }
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
        zip.generateNodeStream({
            type: 'nodebuffer',
            streamFiles: true
        })
            .pipe(stream)
            .on('finish', () => {
                resolve();
            });
    });
}
