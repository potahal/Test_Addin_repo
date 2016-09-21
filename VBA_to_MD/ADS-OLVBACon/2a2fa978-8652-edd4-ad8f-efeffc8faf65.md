

---
ms.Toctitle:受信トレイで件名に"Office"を含むアイテムを検索します。
title:受信トレイで件名に"Office"を含むアイテムを検索します。
ms.ContentId:2a2fa978-8652-edd4-ad8f-efeffc8faf65
---
# 受信トレイで件名に"Office"を含むアイテムを検索します。




このトピックでは、DASL クエリを使用して、受信トレイのアイテムのうちで件名に "Office" を含むものを検索する 2 つのコード サンプルを示します。1 つ目のコード サンプルでは **Folder.GetTable**、2 つ目のコード サンプルでは **Application.AdvancedSearch** を使用して、DASL クエリを適用します。



各コード サンプルでは、プロパティ **http://schemas.microsoft.com/mapi/proptag/0x0037001E** (MAPI ID 名前空間により参照される **Subject** プロパティ) の DASL フィルターでコンテンツ インデックスのキーワード **ci_phrasematch** を使用して、単語 "office" を件名から検索します。 これは、受信トレイ内のアイテムにフィルターを適用し (**Folder.GetTable** または **Application.AdvancedSearch** を使用します)、検索から返された各アイテムの件名を出力します。

>[!NOTE]
>一致では大文字と小文字が区別されませんので、Folder.GetTable または Application.AdvancedSearch では、件名に "Office" または "office" を含むすべてのアイテムが返されます。各サンプルは、出力結果の Table 内の各行の件名を出力します。パフォーマンスを向上させるには、Search.Results オブジェクトではなく、軽量の Table オブジェクトの使用を選択します。Subject プロパティは、どのフォルダーの検索で返される Table にも含まれています。しかし、Outlook の他のフォルダーのように、受信トレイにはさまざまな種類のアイテムを持つことができ、単一の種類のアイテムに限定されていません。受信トレイの特定の種類のアイテムに固有のプロパティにアクセスする場合は、Columns.Add を使用してそのプロパティを取り込み、Table を更新します。そして Table で返される各行について、アイテムのメッセージの種類を確認したうえでプロパティにアクセスします。





次のコード サンプルでは、**Folder.GetTable** を使用して検索を実行します。

```vba
Sub RestrictTableForInbox() 
    Dim oT As Outlook.Table 
    Dim strFilter As String 
    Dim oRow As Outlook.Row 
     
    'Construct filter for Subject containing 'Office' 
    Const PropTag  As String = "http://schemas.microsoft.com/mapi/proptag/" 
    strFilter = "@SQL=" & Chr(34) & PropTag  _ 
        & "0x0037001E" & Chr(34) & " ci_phrasematch 'Office'" 
     
    'Do search and obtain Table on Inbox 
    Set oT = Application.Session.GetDefaultFolder(olFolderInbox).GetTable(strFilter) 
     
    'Print Subject of each returned item 
    Do Until oT.EndOfTable 
        Set oRow = oT.GetNextRow 
        Debug.Print oRow("Subject") 
    Loop 
End Sub
```




次のコード サンプルでは **Application.AdvancedSearch** を使用して検索を実行します。

```vba
Public blnSearchComp As Boolean 
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
    MsgBox "The AdvancedSearchComplete Event fired" 
    blnSearchComp = True 
End Sub 
 
Sub TestSearchWithTable() 
    Dim oSearch As Search 
    Dim oTable As Table 
    Dim strQuery As String 
    Dim oRow As Row 
         
    blnSearchComp = False 
     
    'Construct filter. 0x0037001E represents Subject 
    strQuery = _ 
        "http://schemas.microsoft.com/mapi/proptag/0x0037001E" & _ 
        " ci_phrasematch 'Office'" 
     
    'Do search 
    Set oSearch = _ 
        Application.AdvancedSearch("Inbox", strQuery, False, "Test") 
    While blnSearchComp = False 
        DoEvents 
    Wend 
 
    'Obtain Table 
    Set oTable = oSearch.GetTable 
     
    'Print Subject of each returned item 
    Do Until oTable.EndOfTable 
        Set oRow = oTable.GetNextRow 
        Debug.Print oRow("Subject") 
    Loop 
End Sub

```



