

---
ms.Toctitle:ServerPolicy オブジェクト (Office)
title:ServerPolicy オブジェクト (Office)
ms.ContentId:ce2a63d2-5deb-b94b-45d7-ed84e9be7deb
---
# ServerPolicy オブジェクト (Office)




Microsoft Office SharePoint Server を実行中のサーバーに格納されたドキュメントの種類に対して指定されたポリシーを表します。

## 注釈
**ServerPolicy**オブジェクトは、作業中の文書の個々 のポリシー定義を表す個別の**PolicyItem**オブジェクトで構成されます。



## 例
次の例では、アクティブなドキュメントに対するすべてのポリシー項目の名前と説明を一覧表示します。

```vba
Sub ListPolicyItems() 
Dim objSrvPolicy As ServerPolicy 
Dim objPolicyItem As PolicyItem 
Dim strPolicyItemList As String 
 
Set objSrvPolicy = ActiveDocument.ServerPolicy 
 
For Each objPolicyItem In objSrvPolicy 
 strPolicyItemList = "Policy Item " & objPolicyItem.Name & " - " & _ 
 objPolicyItem.Description & vbCrLf 
Next 
MsgBox (strPolicyItemList) 
 
End Sub 

```




## Related Topics

[オブジェクト モデル リファレンス](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)

[ServerPolicy オブジェクトのメンバー](ed14d9a8-6159-f175-9078-181331ebfb03.md)




