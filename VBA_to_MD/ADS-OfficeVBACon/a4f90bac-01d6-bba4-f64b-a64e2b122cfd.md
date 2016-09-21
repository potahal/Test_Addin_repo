

---
ms.Toctitle:CustomXMLPart オブジェクト (Office)
title:CustomXMLPart オブジェクト (Office)
ms.ContentId:a4f90bac-01d6-bba4-f64b-a64e2b122cfd
---
# CustomXMLPart オブジェクト (Office)




**CustomXMLParts** コレクション内の単一の **CustomXMLPart** を表します。

>[!NOTE]
>
              UNRESOLVED_TOKEN_VAL(osoffdevnodtdsincustomxml)
            





## 例
次の例では、 **CustomXMLPart**オブジェクトに部品を追加します。

```vba
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```




## Related Topics

[オブジェクト モデル リファレンス](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)

[CustomXMLPart オブジェクトのメンバー](76fe85f4-5a35-7d12-2989-6f17a094dcdf.md)




