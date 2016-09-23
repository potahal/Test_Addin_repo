
# ShapeRange.ZOrderPosition プロパティ (プロジェクト)
図形範囲の z オーダーでの位置を取得します。読み取り専用 **Long**です。

## 構文

 _式_. **ZOrderPosition**

 _式_ ShapeRange **ShapeRange** オブジェクトを表す変数。


## 注釈

Z オーダーで図形の位置を設定するには、 [ZOrder](e8badff9-fbe5-b6b8-8c33-68cfde3bef38.md)メソッドを使用します。

図形の z オーダーでの位置は、  **Shapes**コレクション内の図形のインデックス番号に対応します。 `myReport`レポート オブジェクトに 4 つの図形がある場合は、式 `myReport.Shapes(1)`は、z オーダーの背面にある図形を取得などにある式 `myReport.Shapes(4)`は、z オーダーの前面にある図形を返します。

 **Shapes**コレクションに図形を追加すると、既定では、z オーダーの前面に図形が追加されます。


## プロパティ値

 **INT**


## 関連項目


#### その他の技術情報


[ShapeRange オブジェクト](315031aa-4b8c-424b-26e7-ce15897beb05.md)