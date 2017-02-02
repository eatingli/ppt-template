# PPT Template

## 簡介
專為範本功能設計的PPT文件操作API，依賴NodeJS。藉著複製投影片(Slide)並代換其文字內容來實作範本功能，不支援增改圖像、表格等元件操作......。

## 名詞解釋
- 簡報(Presentation)，代表整個PPT文件。
- 投影片(Slide)，代表簡報中的其中一頁。

## 建議用法
1. 製作範本PPT文件，包含美術、排版等。
2. 欲替換的文字用自訂的字串來佔位，建議使用中括號與有意義字詞來表示，例如：[Title]。
3. 用ppt-template API讀取PPT檔案。
4. 讀取並複製投影片。
5. 用實際內容來取代原本的佔位字串。
6. 將完成的投影片加入陣列中，按照想要的順序排序。
7. 產生簡報。
8. 輸出PPT檔案。

## APIs

- 讀取PPT檔案
```
    //從串流讀取
    myPresentation.load(...)

    //從檔案讀取
    myPresentation.loadFile(...)
```

- 讀取投影片數量
```
    myPresentation.getSlideCount()
```

- 讀取投影片
```
    myPresentation.getSlide(slideIndex)
```

- 產生簡報
```
    myPresentation.generate(newSlides)
```

- 輸出PPT檔案
```
    //輸出成檔案
    newPresentation.saveAs(...)
        
    //從串流輸出
    newPresentation.streamAs(...)
```

- 複製投影片
```
    mySlide.clone()
```

- 投影片文字取代
```
    mySlide.fill()
```


## 完整範例
```
    var PPT_Template = require('ppt-template');
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

        //投影片填值
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
```

## 功能
- 讀取Microsoft Office的.pptx檔案，解壓縮後和相關操作方法包裝成簡報物件(Presentation)。
- 利用Presentation物件提供的方法getSlide()，取得樣板投影片物件(Slide)。
- 利用Slide物件提供的複製方法clone()，拷貝Slide物件。
- 利用Slide物件提供的代換內容方法fill()，填入實際內容。
- 利用Presentation物件提供的產生方法generate()，將完成操作Slide集合，製作成新的.pptx檔案。
- 利用Presentation物件提供的輸出方法streamAs()、saveAs()，自訂串流或另外新檔

## 指令
- 下載相依模組
```
    npm install
```
- 建立 
``` 
    npm run build
```
- 測試
```
    npm run test 
```

## 其他
- 引入Promise。