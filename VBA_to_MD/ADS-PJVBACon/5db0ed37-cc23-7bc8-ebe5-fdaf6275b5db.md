

---
ms.Toctitle:Months オブジェクト (Project)
title:Months オブジェクト (Project)
ms.ContentId:5db0ed37-cc23-7bc8-ebe5-fdaf6275b5db
---
# Months オブジェクト (Project)




[Month](5ee32f12-72aa-fa16-ead2-97949005cd7c.md) オブジェクトのコレクションを格納します。

## 注釈
1 つの **Month** オブジェクトを取得するには、**Months**(*Index*) を使用します (引数 *Index* には月のインデックス番号、月名、または **PjMonth** クラスの定数を指定)。



## 例
**Months コレクション オブジェクトの使い方**



次の例では、選択された各リソースの 2012年の各月における稼働日の数をカウントします。

```sourcecode
Dim R As Resource 
Dim D As Integer, M As Integer, WorkingDays As Integer 
 
For Each R In ActiveSelection.Resources() 
    WorkingDays = 0 

    With R.Calendar.Years(2012) 
        For M = 1 To .Months.Count 
            WorkingDays = 0 
            For D = 1 To .Months(M).Days.Count 
                If .Months(M).Days(D).Working = True Then 
                    WorkingDays = WorkingDays + 1 
                End If 
            Next D 

            MsgBox "There are " & WorkingDays & " working days in " & _
                .Months(M).Name & " for " & R.Name & "." 
        Next M 
    End With 
Next R
```




**Months コレクションの使い方**



**月**コレクションを取得するのにには、**数か月**のプロパティを使用します。次の例では、2012 年 5年月の数をカウントします。

```vba
ActiveProject.Calendar.Years(2012).Months.Count
```




## Related Topics

[プロジェクト オブジェクト モデル](900b167b-88ec-ea88-15b7-27bb90c22ac6.md)




