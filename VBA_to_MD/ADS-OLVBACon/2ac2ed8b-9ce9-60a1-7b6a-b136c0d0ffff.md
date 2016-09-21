

---
ms.Toctitle:CardView.Filter プロパティ (Outlook)(機械翻訳)
title:CardView.Filter プロパティ (Outlook)(機械翻訳)
ms.ContentId:2ac2ed8b-9ce9-60a1-7b6a-b136c0d0ffff
---
# CardView.Filter プロパティ (Outlook)(機械翻訳)




返すまたは、ビューのフィルターを表す**文字列**値を設定します。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Filter**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **CardView** オブジェクトを表す変数を指定します。



## 注釈
このプロパティの値は、DAV Searching and Locating (DASL) 構文の文字列で、ビューの現在のフィルターを表します。DASL 構文を使用してビューのアイテムをフィルターにかける方法の詳細については、「[アイテムをフィルターにかける](4038e042-1b07-5d18-18b0-c2b58c9c42da.md)」を参照してください。



## 例
次の Visual Basic for Applications (VBA) の例は、 **Explorer**オブジェクトでは、次に、過去 1 週間に受信した Outlook アイテムのみを表示する**ビュー**の**フィルター**のプロパティがオブジェクトの**示します**プロパティを使用して**View**オブジェクトを取得します。

```vba
Private Sub FilterViewToLastWeek() 
 
 Dim objView As View 
 
 
 
 ' Obtain a View object reference to the current view. 
 
 Set objView = Application.ActiveExplorer.CurrentView 
 
 
 
 ' Set a DASL filter string, using a DASL macro, to show 
 
 ' only those items that were received last week. 
 
 objView.Filter = "%lastweek(""urn:schemas:httpmail:datereceived"")%" 
 
 
 
 ' Save and apply the view. 
 
 objView.Save 
 
 objView.Apply 
 
End Sub 
 

```




## Related Topics

[CardView オブジェクトのメンバー](8b9eda10-1ece-c961-e432-3fca6dfb4f07.md)

[CardView オブジェクト](cdac229b-f2b6-9ecb-e1a7-b53509426570.md)




