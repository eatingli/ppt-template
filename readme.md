# PPT Template

## Introduce

### PPT Template + Customized Content = Generate Result！

![template](/img/01.png)
![customized content](/img/02.png)
![generate result](/img/03.png)

## Dictionary

- **Presentation** - Whole PPT document
- **Slide** - A page of presentation

## Use

1. Prepare the PPT document as template.
2. Put **place-holder** text at variable area. Recommend use meaning string surrounding by brackets.

    eg. **[Title]**、**[Main Content]**

3. Load pptx file by ppt-template API.
4. Get and clone template slide, then replace variable by customized content.
5. Put new sildes in array with wanna order, then generate presentation document.
6. Output your pptx file.

## APIs

- **Load PPT Document**

        // From stream
        myPresentation.load(...)

        // From file
        myPresentation.loadFile(...)


- **Get Silde Count**

        myPresentation.getSlideCount()


- **Get Slide by Index (Base from index 1)**

        myPresentation.getSlide(slideIndex)


- **Generate PPT Document**

        myPresentation.generate(newSlides)

- **Output pptx**

        // Output file
        newPresentation.saveAs(...)

        // Output stream
        newPresentation.streamAs(...)


- **Clone Slide**

        mySlide.clone()


- **Fill Content**

        pair = {key:'place-holder', value:'content'}
        mySlide.fill(pair)

        pairs = [pair, pair]
        mySlide.fillAll(pairs)



## Example

### [Example Code](/example/example.js)


## Command

- **Initial**

        npm install

- **Babel Build**

        npm run build

- **Mocha Test**

        npm run test

- **Example**

        npm run example