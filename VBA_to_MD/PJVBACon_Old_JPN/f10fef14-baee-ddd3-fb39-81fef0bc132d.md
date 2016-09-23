
# Shapes.Value プロパティ (プロジェクト)
 **Shapes**コレクション内の個々 の **Shape**オブジェクトを取得します。読み取り専用 **Shape**です。

## 構文

 _式_. **Value**

 _式_ **Shapes** オブジェクトを表す変数。


### パラメーター



|**名前**|**必須/オプション**|**データ型**|**説明**|
|:-----|:-----|:-----|:-----|
| _Index_|必須|**Variant**|図形の名前を **String**の値、または図形のインデックス番号を序数の **Long**値を指定できます。|

## 注釈

 **Value**は、  **Shapes**オブジェクトの既定のプロパティです。たとえば、という名前のテーブルを含むテーブルのテストレポートを作成します。VBE の **イミディ エイト**ウィンドウで次のステートメントは、テーブルの名前を印刷します。


```
? ActiveProject.Reports("Table Tests").Shapes.Value(1).Name
```

 **Shapes**プロパティを省略して、次のステートメントは効果的に前のステートメントと同じです。




```
? ActiveProject.Reports("Table Tests").Shapes(1).Name
```

 **Shapes.Item** **Shapes.Value**、  **Item**メソッドでは同じような役割を果たします。




```
? ActiveProject.Reports("Table Tests").Shapes.Item(1).Name
```


## プロパティ値

 **SHAPE**


## 関連項目


#### その他の技術情報


[図形オブジェクト](6e42040c-dd5a-de4c-afa8-f9e33d1e5054.md)
[Item メソッド](43fba4f4-f3d3-20a0-2c77-15e31dcdcbf5.md)