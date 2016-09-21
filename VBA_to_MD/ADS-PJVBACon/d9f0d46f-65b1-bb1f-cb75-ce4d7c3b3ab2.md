
---
ms.Toctitle:ShapeRange.ZOrderPosition プロパティ (プロジェクト)
title:ShapeRange.ZOrderPosition プロパティ (プロジェクト)
ms.ContentId:d9f0d46f-65b1-bb1f-cb75-ce4d7c3b3ab2
---
# ShapeRange.ZOrderPosition プロパティ (プロジェクト)





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**ZOrderPosition**




            UNRESOLVED_TOKEN_VAL(offexpression)ShapeRange **ShapeRange** オブジェクトを表す変数。



## 注釈
Z オーダーで図形の位置を設定するには、 [ZOrder](e8badff9-fbe5-b6b8-8c33-68cfde3bef38.md)メソッドを使用します。



図形の z オーダーでの位置は、 **Shapes**コレクション内の図形のインデックス番号に対応します。`myReport`レポート オブジェクトに 4 つの図形がある場合は、式`myReport.Shapes(1)`は、z オーダーの背面にある図形を取得などにある式`myReport.Shapes(4)`は、z オーダーの前面にある図形を返します。



**Shapes**コレクションに図形を追加すると、既定では、z オーダーの前面に図形が追加されます。



## プロパティ値
**INT**



## Related Topics

[ShapeRange オブジェクト](315031aa-4b8c-424b-26e7-ce15897beb05.md)




