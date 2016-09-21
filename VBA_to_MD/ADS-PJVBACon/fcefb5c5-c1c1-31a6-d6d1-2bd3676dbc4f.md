

---
ms.Toctitle:Day.Shift5 プロパティ (Project)
title:Day.Shift5 プロパティ (Project)
ms.ContentId:fcefb5c5-c1c1-31a6-d6d1-2bd3676dbc4f
---
# Day.Shift5 プロパティ (Project)




1 日に 5 番目の勤務シフトを表す**Shift**オブジェクトを取得します。読み取り専用で**shift キーを押し**します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Shift5**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Day** オブジェクトを表す変数です。



## 例
次の使用例は、金曜日の作業時間を半日にします。1 番目の稼働時間帯を午前 8 時から正午までに設定し、他の稼働時間帯の値はクリアします。

```vba
Sub HalfDayFridays() 
 
 With ActiveProject.Calendar.WeekDays(pjFriday) 
 .Shift1.Start = #8:00:00 AM# 
 .Shift1.Finish = #12:00:00 PM# 
 .Shift2.Clear 
 .Shift3.Clear 
 End With 
 
End Sub
```





