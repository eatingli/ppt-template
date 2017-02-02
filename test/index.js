var PPT_Template = require('../');

var Presentation = PPT_Template.Presentation;


var myPresentation = new Presentation();


myPresentation.loadFile('test/test.pptx')

.then(() => {
	console.log('Read Presentation File Successfully!');
})

.then(() => {


	var slideCount = myPresentation.getSlideCount();
	console.log('Slides Count is ', slideCount);

	var slideIndex1 = 1;
	var slideIndex2 = 1;
	var slideIndex3 = 2;

	var cloneSlide1, cloneSlide2, cloneSlide3;

	if(slideIndex1 <= slideCount && slideIndex2 <= slideCount && slideIndex3 <= slideCount){
		
		cloneSlide1 = myPresentation.getSlide(slideIndex1).clone();
		cloneSlide2 = myPresentation.getSlide(slideIndex2).clone();
		cloneSlide3 = myPresentation.getSlide(slideIndex3).clone();

		console.log('Editing Slide...');
	}else{
		console.log('Slide Does Not Exist');
	}

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

	var newSlides = [cloneSlide1, cloneSlide2, cloneSlide3];
	return myPresentation.generate(newSlides);
})

.then((newPresentation) => {
	console.log('Generate New Presentation Successfully');

	return newPresentation.saveAs('test/output.pptx');
})

.then(() => {
	console.log('Save Successfully');
})

.catch((err) => {
	console.error(err);
});