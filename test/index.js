
var PPT_Template = require('../');
var Presentation = PPT_Template.Presentation;

//建立物件
var myPresentation = new Presentation();

//讀取.pptx檔案，接下來透過Promise操作。
myPresentation.loadFile('test/test.pptx')

.then(() => {
	console.log('Read Presentation File Successfully!');
})

.then(() => {

	//讀取投影片數量
	var slideCount = myPresentation.getSlideCount();
	console.log('Slides Count is ', slideCount);

	//透過索引來取得對應投影片，第一張投影片索引為1
	var slideIndex1 = 1;
	var slideIndex2 = 1;
	var slideIndex3 = 2;

	//宣告投影片變數
	var cloneSlide1, cloneSlide2, cloneSlide3;

	//檢查投影片索引
	if(slideIndex1 <= slideCount && slideIndex2 <= slideCount && slideIndex3 <= slideCount){
		
		//取得並複製投影片
		cloneSlide1 = myPresentation.getSlide(slideIndex1).clone();
		cloneSlide2 = myPresentation.getSlide(slideIndex2).clone();
		cloneSlide3 = myPresentation.getSlide(slideIndex3).clone();

		console.log('Editing Slide...');
	}else{
		console.log('Slide Does Not Exist');
	}

	//投影片填值
	cloneSlide1.fill([{
            key: '[Title]',
            value: 'Hello PPT'
        }, {
            key: '[Title2]',
            value: 'this is a sample'
        }, {
            key: '[Description]',
            value: '~~~*^@#%(^(!#~'
        }]);

	cloneSlide3.fill([{
            key: '[Content1]',
            value: 'content~~~~'
        }, {
            key: '[Content2]',
            value: 'little content~~~~~~'
        }]);

	//將處理好的投影片組織到陣列中，產生新的簡報物件
	var newSlides = [cloneSlide1, cloneSlide2, cloneSlide3];
	return myPresentation.generate(newSlides);
})

.then((newPresentation) => {

	console.log('Generate New Presentation Successfully');

	//輸出簡報檔案
	return newPresentation.saveAs('test/output.pptx');
})

.then(() => {
	console.log('Save Successfully');
})

.catch((err) => {
	console.error(err);
});