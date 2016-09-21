

---
ms.Toctitle:WeekDays.Count プロパティ (Project)
title:WeekDays.Count プロパティ (Project)
ms.ContentId:6343346c-dbfc-b36b-eaf4-ddcc2e6f745d
---
# WeekDays.Count プロパティ (Project)




**平日**のコレクション内の項目数を取得します。 読み取り専用**の整数**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Count**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **WeekDays** オブジェクトを表す変数を指定します。



## 例
次の使用例は、指定したリソースのカレンダーの 1 週間が 7 日間であることを示します。

```vba
Debug.Print ActiveProject.Resources(1).Calendar.WorkWeeks(1).WeekDays.Count
```




## Related Topics

[WeekDays コレクション オブジェクト](757437a0-e2ff-0027-f044-87d1cb357f62.md)




