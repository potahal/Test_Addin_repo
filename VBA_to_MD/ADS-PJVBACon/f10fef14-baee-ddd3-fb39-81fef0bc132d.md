
---
ms.Toctitle:Shapes.Value プロパティ (プロジェクト)
title:Shapes.Value プロパティ (プロジェクト)
ms.ContentId:f10fef14-baee-ddd3-fb39-81fef0bc132d
---
# Shapes.Value プロパティ (プロジェクト)





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Value**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Shapes** オブジェクトを表す変数。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Index*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**Variant**|図形の名前を**String**の値、または図形のインデックス番号を序数の**Long**値を指定できます。|





## 注釈
**Value**は、 **Shapes**オブジェクトの既定のプロパティです。たとえば、という名前のテーブルを含む**テーブルのテスト**レポートを作成します。VBE の**イミディ エイト**ウィンドウで次のステートメントは、テーブルの名前を印刷します。

```vba
? ActiveProject.Reports("Table Tests").Shapes.Value(1).Name
```




**Shapes**プロパティを省略して、次のステートメントは効果的に前のステートメントと同じです。

```vba
? ActiveProject.Reports("Table Tests").Shapes(1).Name
```




**Shapes.Item****Shapes.Value**、 **Item**メソッドでは同じような役割を果たします。

```vba
? ActiveProject.Reports("Table Tests").Shapes.Item(1).Name
```




## プロパティ値
**SHAPE**



## Related Topics

[図形オブジェクト](6e42040c-dd5a-de4c-afa8-f9e33d1e5054.md)

[Item メソッド](43fba4f4-f3d3-20a0-2c77-15e31dcdcbf5.md)




