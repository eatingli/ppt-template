
import entities from 'entities'


export default class Slide {

	constructor(rel, content) {

		// ppt/slides/_rels/slide(i).xml.rels
		this.rel = rel;

		// ppt/slides/slide(i).xml
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
	fill(pair) {

		// 檢查key 和value是否存在

		// 處理 XML Entities
		let value = entities.encodeXML(pair.value);
		let key = pair.key;

		// offset: 避免遞迴置換
		let offset = 0;
		let temp = 0;

		// Replace All
		while ((temp = this.content.indexOf(key, offset)) > -1) {

			this.content = Slide.replace(this.content, offset, key, value);
			offset = temp + value.length;
		}
	}

	/**
	 * 
	 */
	fillAll(pairs) {
		pairs.forEach((pair) => {
			this.fill(pair)
		});
	}

	/**
	 * 
	 */
	static replace(str, offset, a, b) {
		let index = str.indexOf(a, offset);
		return (index > -1) ? str.substring(0, index) + str.substring(index, str.length).replace(a, b) : str;
	}

	/**
	 * 
	 */
	static pair(key, value) {
		return { key: key, value: value };
	}
}