

---
ms.Toctitle:Task.HyperlinkAddress プロパティ (Project)
title:Task.HyperlinkAddress プロパティ (Project)
ms.ContentId:0fd6c70e-df9e-1d6e-df65-aa1de2f98b44
---
# Task.HyperlinkAddress プロパティ (Project)




ドキュメントの URL または UNC のパスを設定します。値の取得および設定が可能です。文字列型 (**String**) の値を使用します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**HyperlinkAddress**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Task** オブジェクトを表す変数を指定します。



## 例
次の使用例は、作業中のプロジェクトのすべてのタスク (サブプロジェクトのタスクを含む) にハイパーリンクを設定します。

```vba
Sub AddHyperlink() 
 Dim T As Task 
 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 T.Hyperlink = "Microsoft" 
 T.HyperlinkAddress = "http://www.microsoft.com/" 
 End If 
 Next T 
 
End Su
```





