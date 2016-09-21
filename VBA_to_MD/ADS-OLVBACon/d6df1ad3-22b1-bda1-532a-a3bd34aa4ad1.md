

---
ms.Toctitle:SelectNamesDialog.SetDefaultDisplayMode メソッド (Outlook)(機械翻訳)
title:SelectNamesDialog.SetDefaultDisplayMode メソッド (Outlook)(機械翻訳)
ms.ContentId:d6df1ad3-22b1-bda1-532a-a3bd34aa4ad1
---
# SelectNamesDialog.SetDefaultDisplayMode メソッド (Outlook)(機械翻訳)




[**名前の選択**] ダイアログ ボックスの既定の表示モードを設定し、キャプションとボタンのラベルを指定します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**SetDefaultDisplayMode**(**defaultMode**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SelectNamesDialog** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*defaultMode*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**OlDefaultSelectNamesDisplayMode**|**[名前の選択**] ダイアログ ボックスの既定のキャプションとボタンのラベルを決定する**OlDefaultSelectNamesDisplayMode**列挙体の定数です。|





## 注釈
**SetDefaultDisplayMode**は省略可能です。**表示**を呼び出す前に**SetDefaultDisplayMode**を呼び出していないと、既定の表示モードは、 **OlDefaultSelectNamesDisplayMode.olDefaultMail**をされます。表示モードを別の値に設定するには、**表示**メソッドを呼び出す前に**SetDefaultDisplayMode**を呼び出す必要があります。



このメソッドを使用すると、キャプション、[**宛先**] ラベル、[**ＣＣ**] ラベル、および [**ＢＣＣ**] ラベルの値をローカライズするリソース ファイルを使わずに、[**名前の選択**] ダイアログ ボックスを表示できます。**Caption**、**ToLabel**、**CcLabel**、および **BccLabel** に独自の値を設定することにより、既定の動作を無効にすることができます。



**SetDefaultDisplayMode**を呼び出した後は、追加のプロパティ (たとえば、 **NumberOfRecipientSelectors**を**olRecipientSelectors.olToCc**に設定) を設定できます。**[名前の選択**] ダイアログ ボックスは、それ以降の設定を確認してください。



## Related Topics

[SelectNamesDialog オブジェクトのメンバー](0f5546af-f89a-8a8b-ced9-a2d646bf9634.md)

[SelectNamesDialog オブジェクト](1522736a-3cad-9f1c-4da9-b52a3a01731c.md)




