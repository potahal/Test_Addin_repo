

---
ms.Toctitle:SmartDocument.RefreshPane メソッド (Office)
title:SmartDocument.RefreshPane メソッド (Office)
ms.ContentId:c37de2c2-f24a-0db2-fda8-cfe7d0b464fb
---
# SmartDocument.RefreshPane メソッド (Office)




作業中の Microsoft Word 文書または Microsoft Excel ブックの [**ドキュメント アクション**] 作業ウィンドウを更新します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**RefreshPane**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SmartDocument** オブジェクトを表す変数を指定します。



## 注釈
**RefreshPane**メソッドでは、作業中の文書に添付された XML 拡張パックがない場合にエラーが発生します。



## 例
次の使用例は、アクティブな Excel ブックに XML 拡張パックが関連付けられているかどうかを調べ、関連付けられている場合、 そのスマート ドキュメントの [**ドキュメント アクション**] 作業ウィンドウを更新します。

```sourcecode
 Dim objSmartDoc As Office.SmartDocument 
 Set objSmartDoc = ActiveWorkbook.SmartDocument 
 If objSmartDoc.SolutionID > "None" Then 
 objSmartDoc.RefreshPane 
 Else 
 MsgBox "No XML expansion pack attached." 
 End If 

```




## Related Topics

[SmartDocument オブジェクトのメンバー](980de42d-6992-6107-a3fb-33e8c78da202.md)

[SmartDocument オブジェクト](b56a86eb-a031-d50b-905e-ef8b91914d61.md)




