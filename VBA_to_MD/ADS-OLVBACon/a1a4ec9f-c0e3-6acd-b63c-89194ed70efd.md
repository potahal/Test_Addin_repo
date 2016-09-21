

---
ms.Toctitle:Application.AdvancedSearchStopped イベント (Outlook)
title:Application.AdvancedSearchStopped イベント (Outlook)
ms.ContentId:a1a4ec9f-c0e3-6acd-b63c-89194ed70efd
---
# Application.AdvancedSearchStopped イベント (Outlook)




指定した**Search**オブジェクトの**Stop**メソッドが実行されたときに発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**AdvancedSearchStopped**(**SearchObject**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*SearchObject*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**Search**|**ここ**から返された**Search**オブジェクトです。|





## 注釈
このイベントが発生した後、object?s の**検索****結果**コレクションは更新されません。このイベントは、プログラムでのみ発生させることができます。



## 例
次の Visual Basic for Applications (VBA) の例は、**受信トレイ**内で件名が "Test" のアイテムの検索を開始し、直ちに検索を停止します。これにより、`AdvanceSearchStopped` イベント プロシージャが実行されます。このサンプル コードは、`ThisOutlookSession` などのクラス モジュールに配置する必要があります。また、イベント プロシージャが UNRESOLVED_TOKEN_VAL(outlooknv1) によって呼び出されるためには、それより前に `StopSearch()` プロシージャが呼び出される必要があります。

```vba
Sub StopSearch() 
 
 Dim sch As Outlook.Search 
 
 Dim strScope As String 
 
 Dim strFilter As String 
 
 strScope = "Inbox" 
 
 strFilter = "urn:schemas:httpmail:subject = 'Test'" 
 
 Set sch = Application.AdvancedSearch(strScope, strFilter) 
 
 sch.Stop 
 
End Sub 
 
 
 
Private Sub Application_AdvancedSearchStopped(ByVal SearchObject As Search) 
 
 'Inform the user that the search has stopped. 
 
 MsgBox "An AdvancedSearch has been interrupted and stopped. " 
 
End Sub 
 
 
 

```




## Related Topics

[Application オブジェクト](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)

[Application オブジェクト メンバー](3519c89c-2353-85ee-7ddc-62e5dd85a8e7.md)




