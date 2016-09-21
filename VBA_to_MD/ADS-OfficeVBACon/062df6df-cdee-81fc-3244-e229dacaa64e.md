

---
ms.Toctitle:DocumentProperty.LinkToContent プロパティ (Office)
title:DocumentProperty.LinkToContent プロパティ (Office)
ms.ContentId:062df6df-cdee-81fc-3244-e229dacaa64e
---
# DocumentProperty.LinkToContent プロパティ (Office)




カスタム ドキュメント プロパティの値がコンテナー ドキュメントのコンテンツにリンクされている場合は**True**です。**False**値が静的である場合。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**LinkToContent**(**pfLinkRetVal**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **DocumentProperty** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*pfLinkRetVal*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|ドキュメント プロパティがコンテナー ドキュメントにリンクされるかどうかを示します。|





## 注釈
このプロパティは、カスタム ドキュメント プロパティにのみ適用されます。組み込みのドキュメント プロパティは、このプロパティの値は**False**です。



**LinkSource**プロパティを使用すると、指定したプロパティのリンクのソースを設定できます。**LinkSource**プロパティを設定する**場合は True**に**なります**プロパティを設定します。
Excel の場合はなりますが**True**に設定、ブックから[LinkSource](3e3a6ebc-615a-298e-c40f-cbb6d5cf63e3.md )用のアドレスまたは範囲名を指定する必要があります。アドレスまたは範囲名は、複数のセルをカバーしている場合、カスタム ドキュメント プロパティは範囲の左上のセルから値を取得します。



## 例
この例では、カスタム ドキュメント プロパティのリンク状態を表示します。例が動作するには、 **dp**が有効な**DocumentProperty**オブジェクトでなければなりません。

```sourcecode
Sub DisplayLinkStatus(dp As DocumentProperty) 
 Dim stat As String, tf As String 
 If dp.LinkToContent Then 
 tf = "" 
 Else 
 tf = "not " 
 End If 
 stat = "This property is " & tf & "linked" 
 If dp.LinkToContent Then 
 stat = stat + Chr(13) & "The link source is " & dp.LinkSource 
 End If 
 MsgBox stat 
End Sub
```




## Related Topics

[DocumentProperty オブジェクトのメンバー](568da0ff-fa90-150a-06ec-611de886334e.md)

[DocumentProperty オブジェクト](dd54ca3c-e0e2-4816-539a-17c5b4a928b1.md)

[同期オブジェクト](1cb049a0-a803-969a-7923-15ddb8da8f3b.md)

[同期オブジェクトのメンバー](748726bd-83de-425a-5af8-177c34e3a013.md)




