# PPT Template

## 功能
- 讀取Microsoft Office的.pptx檔案，解壓縮後包裝成簡報物件(Presentation物件)，並提供一些操作方法。
- 利用Presentation物件提供的方法getSlide()，取得樣板投影片物件(Slide物件)。
- 利用Slide物件提供的複製方法clone()，拷貝Slide物件。
- 利用Slide物件提供的代換內容方法fill()，填入實際內容。
- 利用Presentation物件提供的產生方法generate()，將完成操作Slide集合，製作成新的.pptx檔案。
- 利用Presentation物件提供的輸出方法streamAs()、saveAs()，自訂串流或另外新檔

## 使用
1.製作簡報範本(PPT Template)，要替換的內容可用 [Content]、[Name]、[Key]...文字來標示。
2.對樣板投影片(Slide)進行複製，並填值。填值格式為map物件的陣列，fill()範例：
```
myCloneSlide.fill([{
            key: '[Content]',
            value: 'My Content'
        }, {
            key: '[Name]',
            value: 'Suwako'
        }])
```
3.填值後產生新的.pptx並另存新檔。

## 開發
- 建立 
``` npm run build ```
- Presention測試
``` npm run test ```

## 其他
- 全面引入Promise。