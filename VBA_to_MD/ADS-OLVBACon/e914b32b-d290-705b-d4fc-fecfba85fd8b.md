

---
ms.Toctitle:テーブル内の複数値を持つプロパティの値にアクセスします。
title:テーブル内の複数値を持つプロパティの値にアクセスします。
ms.ContentId:e914b32b-d290-705b-d4fc-fecfba85fd8b
---
# テーブル内の複数値を持つプロパティの値にアクセスします。




一般に、複数値を持つプロパティを **Table** に追加するときに、明示的な組み込みの名前を使用すると、プロパティ値の形式はコンマ区切りの文字列になります。複数値を持つプロパティを **Table** に追加するときに、名前空間による参照を使用すると、プロパティ値の形式はバリアントの配列になります。



次のコード サンプルでは、複数値を持つ **Categories** プロパティを、その名前空間 **urn:schemas-microsoft-com:office:office#Keywords** を参照する名前を使用して、**Table** に追加します。**Table** の各行について、**Categories** の値を取得するために、

```sourcecode
oRow("urn:schemas-microsoft-com:office:office#Keywords")
```




variant、およびバリアント型の配列の要素を列挙します。アイテムをするが割り当てられていないことすべてのカテゴリでは、variant に注意してくださいし、バリアント型の配列の要素を列挙します。アイテムのことが割り当てられていないことすべてのカテゴリでは、注

```sourcecode
oRow("urn:schemas-microsoft-com:office:office#Keywords")
```




は Empty 値を返します。

```sourcecode
Sub TableCategories() 
    Dim oT As Outlook.Table 
    Dim oRow As Outlook.Row 
    Dim varCat 
    Dim j As Integer 
    Dim strCategories As String 
 
    Set oT = Application.ActiveExplorer.CurrentFolder.GetTable() 
    oT.Columns.Add ("urn:schemas-microsoft-com:office:office#Keywords") 
    oT.Sort "LastModificationTime", True 
    Do Until oT.EndOfTable 
        Set oRow = oT.GetNextRow 
        'Obtain any values of the Categories property 
        varCat = oRow("urn:schemas-microsoft-com:office:office#Keywords") 
        If Not (IsEmpty(varCat)) Then 
            'Form a string out of the item's categories 
            For j = 0 To UBound(varCat) 
                strCategories = strCategories & (varCat(j)) & ", " 
            Next 
            'Remove last trailing ", " 
            strCategories = Left(strCategories, Len(strCategories) - 2) 
        Else 
            'The item does not have any categories 
            strCategories = "" 
        End If 
        Debug.Print ("Subject: " _ 
           & oRow("Subject") & vbCrLf & "Categories: ") & strCategories & vbCrLf 
    Loop 
End Sub
```



