

---
ms.Toctitle:AppointmentItem.Duration プロパティ (Outlook)(機械翻訳)
title:AppointmentItem.Duration プロパティ (Outlook)(機械翻訳)
ms.ContentId:eea64bdd-c19b-01c7-4fdb-111df86de2c4
---
# AppointmentItem.Duration プロパティ (Outlook)(機械翻訳)




**長い**示す**AppointmentItem**の期間 (分単位) を設定または返します。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Duration**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **AppointmentItem** オブジェクトを表す変数を指定します。



## 例
この Visual Basic for Applications の例では、 **Application.CreateItem**を使用して予定を作成し、必須および任意の出席者に会議出席依頼に変換するには、「会議」にミーティングのステータスを設定するのには**AppointmentItem.MeetingStatus**を使用します。

```vba
Sub ScheduleMeeting() 
 
 Dim myItem as AppointmentItem 
 
 Dim myRequiredAttendee As Recipient 
 
 Dim myOptionalAttendee As Recipient 
 
 Dim myResourceAttendee As Recipient 
 
 
 
 Set myItem = Application.CreateItem(olAppointmentItem) 
 
 myItem.MeetingStatus = olMeeting 
 
 myItem.Subject = "Strategy Meeting" 
 
 myItem.Location = "Conference Room B" 
 
 myItem.Start = #9/24/2002 1:30:00 PM# 
 
 myItem.Duration = 90 
 
 Set myRequiredAttendee = myItem.Recipients.Add ("Nate Sun") 
 
 myRequiredAttendee.Type = olRequired 
 
 Set myOptionalAttendee = myItem.Recipients.Add ("Kevin Kennedy") 
 
 myOptionalAttendee.Type = olOptional 
 
 Set myResourceAttendee = myItem.Recipients.Add("Conference Room B") 
 
 myResourceAttendee.Type = olResource 
 
 myItem.Display 
 
End Sub
```




## Related Topics

[AppointmentItem オブジェクトのメンバー](c72c459d-6d3c-7a05-aa4a-b1b767ddc0b2.md)

[AppointmentItem オブジェクト](204a409d-654e-27aa-643a-8344c631b82d.md)




