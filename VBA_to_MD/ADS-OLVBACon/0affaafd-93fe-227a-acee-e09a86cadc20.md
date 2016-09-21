

---
ms.Toctitle:テーブルのオブジェクト (Outlook)(機械翻訳)
title:テーブルのオブジェクト (Outlook)(機械翻訳)
ms.ContentId:0affaafd-93fe-227a-acee-e09a86cadc20
---
# テーブルのオブジェクト (Outlook)(機械翻訳)




**Folder** オブジェクトまたは **Search** オブジェクトのアイテム データの集合を表します。アイテムはテーブルの行になり、プロパティはテーブルの列になります。

## 注釈
**テーブル**は読み取り専用で動的な行セットの**フォルダー**または**検索**オブジェクト内のデータを表します。**マッチング**または**これら**を使用するには、フォルダーまたは検索フォルダー内の項目のセットを表す**Table**オブジェクトを取得します。**マッチング**から**Table**オブジェクトを取得した場合は、フォルダー内のアイテムのサブセットを取得するのにはさらに ( **Table.Restrict**) でフィルターを指定できます。任意のフィルターを指定しない場合、フォルダー内のすべてのアイテムが表示されます。



既定では、返される**テーブル**内の各項目には、そのプロパティの既定のサブセットのみが含まれています。フォルダー内の項目としては、**テーブル**の各行、各列をし、インメモリの軽量な行セットとなり、高速な列挙では、**テーブル**のプロパティ、フォルダー内のアイテムのフィルター処理と見なすことができます。基になるフォルダーの追加と削除は、**テーブル**内の行に反映されますが、行の削除と追加、変更、**テーブル**上で任意のイベントはサポートされません。オブジェクトは、**テーブル**の行が必要な場合は、既定の**テーブル**内の列のエントリ Id からその行のエントリ ID を取得して、完全なアイテムを取得するのには、**名前空間**オブジェクトの**GetItemFromID**メソッドを使用して、 **MailItem****ContactItem**など、読み取り/書き込み操作をサポートします。



**テーブル**の既定の列の詳細については、 [Table オブジェクトに表示される既定のプロパティ](649c64f3-2d1e-23f1-bf13-3368da79e62b.md)を参照してください。



**テーブル**・ オブジェクトの詳細については、[列挙、検索、およびフォルダー内のアイテムのフィルタ リング](d786d292-7a0e-0e1a-e132-affbfde37744.md)を参照してください。



## 例
次のコード例は、 **Table**オブジェクトでのアイテムの**LastModificationTime**プロパティに基づいてフィルター処理されたセットを返すことができる方法を示しています。既定のプロパティと項目の特定のプロパティを一覧表示する方法も示します。

```vba
Sub DemoTable() 
 
 'Declarations 
 
 Dim Filter As String 
 
 Dim oRow As Outlook.Row 
 
 Dim oTable As Outlook.Table 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 'Get a Folder object for the Inbox 
 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 
 
 'Define Filter to obtain items last modified after May 1, 2005 
 
 Filter = "[LastModificationTime] > '5/1/2005'" 
 
 'Restrict with Filter 
 
 Set oTable = oFolder.GetTable(Filter) 
 
 
 
 'Remove all columns in the default column set 
 
 oTable.Columns.RemoveAll 
 
 'Specify desired properties 
 
 With oTable.Columns 
 
 .Add ("Subject") 
 
 .Add ("LastModificationTime") 
 
 'PR_ATTR_HIDDEN referenced by the MAPI proptag namespace 
 
 .Add ("http://schemas.microsoft.com/mapi/proptag/0x10F4000B") 
 
 End With 
 
 
 
 'Enumerate the table using test for EndOfTable 
 
 Do Until (oTable.EndOfTable) 
 
 Set oRow = oTable.GetNextRow() 
 
 Debug.Print (oRow("Subject")) 
 
 Debug.Print (oRow("LastModificationTime")) 
 
 Debug.Print (oRow("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) 
 
 Loop 
 
End Sub
```




## Related Topics

[テーブル オブジェクトのメンバー](bd9db35d-0738-22cf-a936-425d5a0ead87.md)

[Outlook オブジェクト モデル リファレンス](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)




