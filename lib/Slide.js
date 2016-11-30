var Slide = module.exports = function(rel, content) {
	this.rel = rel;
	this.content = content;
}

Slide.prototype.clone = function() {
	return new Slide(this.rel, this.content);
}

Slide.prototype.fill = function(data) {

	var self = this;

	data.forEach(function(d) {

		var index = 0;
		var temp = 0;

		while ((temp = self.content.indexOf(d.key, index)) > -1) {
			self.content = replace(self.content, index, d.key, d.value)
			index = temp + d.value.length;
		}
	});

	return this;
}

function replace(str, index, a, b) {
	var index = str.indexOf(a, index);
	if (index > -1) {
		return str.substring(0, index) + str.substring(index, str.length).replace(a, b);
	}
}

//var test = "AAABBBCCCAAABBBCCCDDD";
//console.log(replace(test, 3, 'A', 'XX'));