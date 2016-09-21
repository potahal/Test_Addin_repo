

---
ms.Toctitle:CommandButton.Picture プロパティ (Outlook フォーム スクリプト)
title:CommandButton.Picture プロパティ (Outlook フォーム スクリプト)
ms.ContentId:b92228be-dda7-fdde-2d0c-8e59f544d8db
---
# CommandButton.Picture プロパティ (Outlook フォーム スクリプト)




コントロールに表示するビットマップのフルパス名を指定する**文字列**を返します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Picture**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **CommandButton** オブジェクトを表す変数です。



## 注釈
**Picture**プロパティにビットマップを割り当てるには、コントロールのプロパティ ページを使用する必要があります。**画像**にビットマップを割り当てるには、Visual Basic の**LoadPicture**関数を使うことはできません。



コントロールに割り当てられているピクチャを削除するには、プロパティ ページで**Picture**プロパティの値をクリックして、 **DEL**キーを押します。**Backspace キー**を押すと、画像は削除されません。



キャプション付きのコントロールでは、**PicturePosition** プロパティを使用して、ピクチャを表示する位置を指定できます。



透明なピクチャは、かすんで表示されることがあります。このような問題を避けるには、ピクチャを不透明に表示できるコントロール上にピクチャを表示します。**Image** は、ピクチャを不透明に表示できます。




