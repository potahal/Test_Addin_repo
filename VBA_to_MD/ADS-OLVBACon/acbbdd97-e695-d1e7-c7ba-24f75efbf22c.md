

---
ms.Toctitle:ColumnFormat オブジェクト (Outlook)(機械翻訳)
title:ColumnFormat オブジェクト (Outlook)(機械翻訳)
ms.ContentId:acbbdd97-e695-d1e7-c7ba-24f75efbf22c
---
# ColumnFormat オブジェクト (Outlook)(機械翻訳)




ビュー内の順序フィールドまたはビュー フィールドの表示プロパティを表します。

## 注釈
**ColumnFormat**オブジェクトは、 **OrderField**オブジェクトまたは**ViewField**オブジェクトの配置やフィールドの種類など、画面のプロパティを表します。ビュー フィールドの表示プロパティにアクセスするのにには、 **ViewField**オブジェクトの**ColumnFormat**プロパティを使用します。



**Label** プロパティを使用すると、フィールドのラベルに使用するテキストを取得または変更でき、**Align** プロパティを使用すると、フィールド内の内容の配置を確認できます。



**FieldType** プロパティを使用すると、そのフィールドに表示されるデータの型と形式を確認でき、**FieldFormat** プロパティを使用すると、そのフィールドのデータの書式設定方法を確認できます。



## 例
次の Visual Basic for Applications (VBA) の例は、コレクション内のラベルと各**ViewField**オブジェクトの XML スキーマの名前を表示する、現在の**TableView**オブジェクトの**ViewFields**コレクションを反復処理します。

```vba
Private Sub DisplayTableViewFields() 
 
 Dim objTableView As TableView 
 
 Dim objViewField As ViewField 
 
 Dim strOutput As String 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTableView Then 
 
 
 
 ' Obtain a TableView object reference for the 
 
 ' current table view. 
 
 Set objTableView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Iterate through the ViewFields collection for 
 
 ' the table view, obtaining the label and the 
 
 ' XML schema name for each field included in 
 
 ' the view. 
 
 For Each objViewField In objTableView.ViewFields 
 
 With objViewField 
 
 strOutput = strOutput & .ColumnFormat.Label & _ 
 
 " (" & .ViewXMLSchemaName & ")" & vbCrLf 
 
 End With 
 
 Next 
 
 
 
 ' Display a dialog box containing the concatenated 
 
 ' view field information. 
 
 MsgBox strOutput 
 
 End If 
 
End Sub 
 

```




## Related Topics

[ColumnFormat オブジェクトのメンバー](7159f452-7a05-f3a3-53f8-0b3f5463d313.md)

[Outlook オブジェクト モデル リファレンス](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)




