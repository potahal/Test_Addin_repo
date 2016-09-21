

---
ms.Toctitle:RecurrencePattern.RecurrenceType プロパティ (Outlook)(機械翻訳)
title:RecurrencePattern.RecurrenceType プロパティ (Outlook)(機械翻訳)
ms.ContentId:bc9b35b5-ef00-e5cf-09cc-ee8743efddcf
---
# RecurrencePattern.RecurrenceType プロパティ (Outlook)(機械翻訳)




定期的なパターンの繰り返しの周期を示す **OlRecurrenceType** クラスの定数を設定します。値の取得および設定が可能です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**RecurrenceType**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **RecurrencePattern** オブジェクトを表す変数を指定します。



## 注釈
**RecurrencePattern**オブジェクトの他のプロパティを設定する前に、 **RecurrenceType**プロパティを設定する必要があります。**RecurrencePattern**プロパティを後で設定することは、次の表に示すように、 **RecurrenceType**の値とは異なります。

|||
|---|---|
|**OlRecurrenceType**|**有効な RecurrencePattern プロパティ**|
|**olRecursDaily**|**Duration**、**EndTime**、**Interval**、**NoEndDate**、**Occurrences**、**PatternStartDate**、**PatternEndDate**、**StartTime**|
|**olRecursWeekly**|**DayOfWeekMask**、**期間**、**終了時刻**、**間隔**、 **noenddate プロパティ**、**出現する**、 **PatternStartDate**、**年**、**開始時刻**|
|**olRecursMonthly**|**DayOfMonth**、**Duration**、**EndTime**、**Interval**、**NoEndDate**、**Occurrences**、**PatternStartDate**、**PatternEndDate**、**StartTime**|
|**olRecursMonthNth**|**DayOfWeekMask**、**Duration**、**EndTime**、**Interval**、**Instance**、**NoEndDate**、**Occurrences**、**PatternStartDate**、**PatternEndDate**、**StartTime**|
|**olRecursYearly**|**DayOfMonth**、**Duration**、**EndTime**、**Interval**、**MonthOfYear**、**NoEndDate**、**Occurrences**、**PatternStartDate**、**PatternEndDate**、**StartTime**|
|**olRecursYearNth**|**DayOfWeekMask**、**Duration**、**EndTime**、**Interval**、**Instance**、**NoEndDate**、**Occurrences**、**PatternStartDate**、**PatternEndDate**、**StartTime**|



## 例
この Visual Basic for Applications の例は、新しく作成された**AppointmentItem**の**RecurrencePattern**オブジェクトを取得するのに**GetRecurrencePattern**を使用します。プロパティ、 **RecurrenceType** 、 **DayOfWeekMask**、 **MonthOfYear**、**インスタンス**、**出現**、**開始時刻**、**終了時刻**、および**件名**が設定されて、予定が保存され、パターンが表示されます:"6 月の最初の月曜日が発生する効果的な 1 2007/6/6/6/2016 午後 2時 00分から午後 5時 00分まで」。

```vba
Sub RecurringYearNth() 
 
 Dim oAppt As AppointmentItem 
 
 Dim oPattern As RecurrencePattern 
 
 Set oAppt = Application.CreateItem(olAppointmentItem) 
 
 Set oPattern = oAppt.GetRecurrencePattern 
 
 With oPattern 
 
 .RecurrenceType = olRecursYearNth 
 
 .DayOfWeekMask = olMonday 
 
 .MonthOfYear = 6 
 
 .Instance = 1 
 
 .Occurrences = 10 
 
 .Duration = 180 
 
 .PatternStartDate = #6/1/2007# 
 
 .StartTime = #2:00:00 PM# 
 
 .EndTime = #5:00:00 PM# 
 
 End With 
 
 oAppt.Subject = "Recurring YearNth Appointment" 
 
 oAppt.Save 
 
 oAppt.Display 
 
End Sub 
 

```




## Related Topics

[RecurrencePattern オブジェクト](36c098f7-59fb-879a-5173-ed0260d13fa4.md)

[RecurrencePattern オブジェクトのメンバー](d282fdb2-2b6d-983d-fe5f-698113d35f89.md)




