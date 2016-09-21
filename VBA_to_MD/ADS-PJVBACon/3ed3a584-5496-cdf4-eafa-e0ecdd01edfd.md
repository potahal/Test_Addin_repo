

---
ms.Toctitle:TimeScaleValue.Clear メソッド (Project)
title:TimeScaleValue.Clear メソッド (Project)
ms.ContentId:3ed3a584-5496-cdf4-eafa-e0ecdd01edfd
---
# TimeScaleValue.Clear メソッド (Project)




タイムスケール領域にあるデータの値をクリアします。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Clear**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TimeScaleValue** オブジェクトを表す変数です。



## 例
次の使用例は、金曜日の稼働時間帯を半日に設定します。1 番目の稼働時間帯を午前 8 時から正午までに設定し、他の稼働時間帯をクリアします。

```vba
Sub HalfDayFridays() 
 With ActiveProject.Calendar.Weekdays(pjFriday) 
 .Shift1.Start = #8:00:00 AM# 
 .Shift1.Finish = #12:00:00 PM# 
 .Shift2.Clear 
 .Shift3.Clear 
 End With 
End Sub
```





