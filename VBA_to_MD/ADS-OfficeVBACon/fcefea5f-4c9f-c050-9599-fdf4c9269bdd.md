

---
ms.Toctitle:WebPageFont.ProportionalFont プロパティ (Office)
title:WebPageFont.ProportionalFont プロパティ (Office)
ms.ContentId:fcefea5f-4c9f-c050-9599-fdf4c9269bdd
---
# WebPageFont.ProportionalFont プロパティ (Office)




ホスト アプリケーションにプロポーショナル フォントを取得または設定します。値の取得および設定が可能です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**ProportionalFont**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **WebPageFont** オブジェクトを表す変数を指定します。



## 注釈
**ProportionalFont**プロパティを設定すると、ホスト アプリケーションは値の有効性をチェックしません。



## 例
次の使用例は、作業中のアプリケーションで使用される英語/西ヨーロッパ言語/その他のラテン系言語の文字セットの、プロポーショナル フォントとプロポーショナル フォント サイズを設定します。

```sourcecode
Application.DefaultWebOptions. _ 
Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) _ 
.ProportionalFont = "Tahoma" 
Application.DefaultWebOptions. _ 
Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) _ 
.ProportionalFontSize = 14.5
```




## Related Topics

[WebPageFont オブジェクト](daf3c079-520d-68bd-ec02-027776074505.md)

[WebPageFont オブジェクトのメンバー](82843862-c4b8-db92-d9a7-da36908a0b5e.md)




