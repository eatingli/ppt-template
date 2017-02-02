# PPT Template

## 簡介
專為範本功能設計的PPT文件操作API，依賴NodeJS。藉著複製投影片(Slide)並代換其文字內容來實作範本功能，不支援增改圖像、表格等元件操作......。

## 建議用法
1. 製作範本PPT文件，包含美術、排版等。
2. 欲替換的文字用自訂的字串佔位，建議使用中括號框住有意義字詞來代表，例如：[Title]。
3. 用ppt-template API讀取簡報檔案。
    ```
    var Presentation = require('ppt-template').Presentation;
    var myPresentation = new Presentation();
    
    myPresentation.loadFile('test/test.pptx')
    .then(() => {
        console.log('Read Presentation File Successfully!');
    })

    ```
4. 讀取並複製(clone)投影片。
    ```
        var cloneSlide = myPresentation.getSlide(1).clone();
    ```
5. 用實際內容取代(fill)原本的佔位字串。
    ```
        cloneSlide.fill([{
                key: '[Title]',
                value: 'Hello PPT'
            }, {
                key: '[Title2]',
                value: 'this is a sample'
            }, {
                key: '[Description]',
                value: '~~~*^@#%(^(!#~'
            }]);
    ```
6. 將完成的投影片加入陣列中，按照要輸出的順序排序。
    ```
        var newSlides = [cloneSlide1, cloneSlide2, cloneSlide3];
    ```
7. 使用投影片陣列來產生(generate)簡報(Presentation)。
    ```
        return myPresentation.generate(newSlides)
    ```
8. 輸出。
    ```
        return newPresentation.saveAs('test/output.pptx');
    ```


## 功能
- 讀取Microsoft Office的.pptx檔案，解壓縮後包裝成簡報物件(Presentation物件)，並提供一些操作方法。
- 利用Presentation物件提供的方法getSlide()，取得樣板投影片物件(Slide物件)。
- 利用Slide物件提供的複製方法clone()，拷貝Slide物件。
- 利用Slide物件提供的代換內容方法fill()，填入實際內容。
- 利用Presentation物件提供的產生方法generate()，將完成操作Slide集合，製作成新的.pptx檔案。
- 利用Presentation物件提供的輸出方法streamAs()、saveAs()，自訂串流或另外新檔

## 指令
- 建立 
``` npm run build ```
- 測試
``` npm run test ```

## 其他
- 引入Promise。