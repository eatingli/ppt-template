
import entities from 'entities'


export default class Slide {

	constructor(rel, content) {

		// ppt/slides/_rels/slideI.xml.rels
		this.rel = rel;

		// ppt/slides/slideI.xml
		this.content = content;
	}

	/**
	 * 
	 */
	clone() {
		return new Slide(this.rel, this.content);
	}

	/**
	 * 
	 */
	fill(word) {

		// 把 "&" 之類的符號轉換 &amp;  (XML Entities)
		let value = entities.encodeXML(word.value);
		let key = word.key;

		// offset: 避免遞迴置換...
		let offset = 0;
		let temp = 0;

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
	fillAll(words) {
		words.forEach((word) => {
			this.fill(word)
		});
	}

}


function replace(str, offset, a, b) {
	let index = str.indexOf(a, offset);
	return (index > -1) ? str.substring(0, index) + str.substring(index, str.length).replace(a, b) : str;
}


// var test = "AABBCCAABB";
// console.log(replace(test, 3, 'A', 'XX')); // ---> AABBCCXXABB
// console.log(replace(test, 3, 'D', 'XX')); // ---> AABBCCAABB