

---
ms.Toctitle:PolicyItem オブジェクト (Office)
title:PolicyItem オブジェクト (Office)
ms.ContentId:aced7bdc-8ef7-2621-f188-f3c1d44ab6dc
---
# PolicyItem オブジェクト (Office)




1 つのポリシーの設定を含む**ServerPolicy**オブジェクト内の項目を表します。

## 注釈
ポリシー項目は、ポリシーの範囲外では存在できません。ポリシー項目は、Microsoft Office SharePoint Server を実行中のサーバーに格納されたドキュメントに対して定義された個別の条件です。



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

[PolicyItem オブジェクトのメンバー](a2e43e08-64bb-f052-78a2-0618e2df46fc.md)




