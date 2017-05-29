
const TEMPLATE = 'example/example.pptx';
const OUTPUT = 'example/output.pptx';

var PPT_Template = require('../');
var Presentation = PPT_Template.Presentation;
var Slide = PPT_Template.Slide;

// Presentation Object
var myPresentation = new Presentation();

console.log('# Load test.pptx as template, then build output.pptx with custom content.');

// Load example.pptx
myPresentation.loadFile(TEMPLATE)

	.then(() => {
		console.log('- Read Presentation File Successfully!');
	})

	.then(() => {

		// get slide conut
		var slideCount = myPresentation.getSlideCount();
		console.log('- Slides Count is ', slideCount);

		// Get slide by index. (Base from 1)
		var slideIndex1 = 1;
		var slideIndex2 = 1;
		var slideIndex3 = 2;

		// Get and clone slide. (Watch out index...)
		let cloneSlide1 = myPresentation.getSlide(slideIndex1).clone();
		let cloneSlide2 = myPresentation.getSlide(slideIndex2).clone();
		let cloneSlide3 = myPresentation.getSlide(slideIndex3).clone();

		// Fill all content
		cloneSlide1.fillAll([
			Slide.pair('[Title]', 'Hello PPT'),
			Slide.pair('[Title2]', 'this is a sample'),
			Slide.pair('[Description]', 'fillAll()')
		]);

		// Fill content
		cloneSlide3.fill(Slide.pair('[Content1]', 'fill() 1'));
		cloneSlide3.fill(Slide.pair('[Content2]', 'fill() 2'));

		// Generate new presention by silde array.
		var newSlides = [cloneSlide1, cloneSlide2, cloneSlide3];
		return myPresentation.generate(newSlides);
	})

	.then((newPresentation) => {
		console.log('- Generate Presentation Successfully');
		return newPresentation;
	})

	.then((newPresentation) => {
		// Output .pptx file
		return newPresentation.saveAs(OUTPUT);
	})

	.then(() => {
		console.log('- Save Successfully');
	})

	.catch((err) => {
		console.error(err);
	});