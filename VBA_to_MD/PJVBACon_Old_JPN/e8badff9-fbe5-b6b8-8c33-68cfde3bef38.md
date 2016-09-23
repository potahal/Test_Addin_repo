
# Shape.ZOrder メソッド (プロジェクト)
図形の前面、またはその他の図形の後ろに移動 (つまり、z オーダーでの位置を変更します)。

## 構文

 _式_. **ZOrder** _(ZOrderCmd)_

 _式_ Shape **Shape** オブジェクトを表す変数。


### パラメーター



|**名前**|**必須/オプション**|**データ型**|**説明**|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|必須|**[MsoZOrderCmd](http://msdn.microsoft.com/en-us/library/office/ff861432%28v=office.15%29)**|その他の図形を基準に図形を移動する場所を指定します。|
| _ZOrderCmd_|必須|MSOZORDERCMD||

### 戻り値

 **Nothing**


## 注釈

図形の z オーダー内の現在の位置を決定するのにには、  **ZOrderPosition**プロパティを使用します。


## 関連項目


#### その他の技術情報


[Shape オブジェクト](d2b32bcd-5595-a4a7-9772-feb25fd0103a.md)
[MsoZOrderCmd](http://msdn.microsoft.com/en-us/library/office/ff861432%28v=office.15%29)
[ZOrderPosition プロパティ](d9f0d46f-65b1-bb1f-cb75-ce4d7c3b3ab2.md)