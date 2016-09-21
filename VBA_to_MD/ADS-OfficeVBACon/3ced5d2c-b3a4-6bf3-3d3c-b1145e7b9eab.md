



---
ms.Toctitle:TextRange2.InsertChartField メソッド (Office)
title:TextRange2.InsertChartField メソッド (Office)
ms.ContentId:3ced5d2c-b3a4-6bf3-3d3c-b1145e7b9eab
---
# TextRange2.InsertChartField メソッド (Office)




フィールドをグラフのデータ ラベルの本文に挿入します。



このメソッドは、グラフのデータ ラベルにのみ適用されます。[TextRange2](a6a59c9b-9b64-c1e2-2e98-a1f99025c877.md)オブジェクトの他の種類でこのメソッドを呼び出すと、実行時エラーが発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**InsertChartField***(ChartFieldType,**Formula,**Position)*




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TextRange2**オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*ChartFieldType*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |[MsoChartFieldType](ce6b367d-d09f-4345-33e3-f181b1a9a41d.md)|グラフのデータ ラベルに挿入するフィールドの種類を指定します。|
|*Formula*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**string**|**MsoChartFieldFormula**定数が*ChartFieldType*パラメーターに渡された場合は、セル (または範囲) を指定します。|
|*Position*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**integer**|グラフのフィールドを挿入する位置の文字位置を指定します。既定では、テキストの末尾にフィールドを追加します。位置の値が範囲外にある場合は、既定値が使用されます。|
|*ChartFieldType*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |MSOCHARTFIELDTYPE||
|*Formula*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |STRING||
|*Position*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |INT||
|名前|必須/オプション|データ型|説明|



### 戻り値
[TextRange2](a6a59c9b-9b64-c1e2-2e98-a1f99025c877.md)






