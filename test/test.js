var assert = require('assert');

var Slide = require('../').Slide;
var Presentation = require('../').Presentation;


describe('# Slide', () => {

    describe('- replace()', () => {

        let testStr = 'AA BB CC AA BB'

        it('回傳成功置換的結果', () => {
            let result = Slide.replace(testStr, 3, 'A', 'X');
            assert.equal(result, 'AA BB CC XA BB');
        });
        it('失敗置換，回傳原字串', () => {
            let result = Slide.replace(testStr, 3, 'D', 'X');
            assert.equal(result, 'AA BB CC AA BB');
        });
    });

    describe('- pair()', () => {
        it('Build key-value pair', () => {
            let pair = Slide.pair('A key', 'A value');
            assert.equal(pair.key, 'A key');
            assert.equal(pair.value, 'A value');
        });
    });

    describe('- clone()', () => {
        let testSlide = new Slide('abc', 'def');
        let cloneOne = testSlide.clone();

        it('複製數值相同的獨立物件', () => {
            assert.equal(cloneOne.rel, testSlide.rel);
            assert.equal(cloneOne.content, testSlide.content);
        });
    });

    describe('- fill()', () => {
        let content = '[title] ~~ [text] ~~ [text]';
        let testSlide = new Slide('', content);

        it('成功代換 pair', () => {
            let pair = Slide.pair('[text]', 'Cat');
            testSlide.fill(pair);
            assert.equal(testSlide.content, '[title] ~~ Cat ~~ Cat');
        });

        it('處理 XML Entities', () => {
            let pair = Slide.pair('[title]', '&');
            testSlide.fill(pair);
            assert.equal(testSlide.content, '&amp; ~~ Cat ~~ Cat');
        });

        it('Pair Value Error', () => {
            assert.throws(() => {
                let pair = { key: 'k' };
                testSlide.fill(pair);
            }, Error);
        });
    });

    describe('- fillAll()', () => {
        let content = '[title] ~~ [text] ~~ [text]';
        let testSlide = new Slide('', content);

        let pairs = [
            Slide.pair('[title]', 'Hello'),
            Slide.pair('[text]', 'Cat')
        ];

        it('成功代換所有 pair', () => {
            testSlide.fillAll(pairs);
            assert.equal(testSlide.content, 'Hello ~~ Cat ~~ Cat');
        });
    });
});