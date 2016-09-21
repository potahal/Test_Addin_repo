

---
ms.Toctitle:Inspector.SetSchedulingStartTime メソッド (Outlook)(機械翻訳)
title:Inspector.SetSchedulingStartTime メソッド (Outlook)(機械翻訳)
ms.ContentId:22e6358a-9dba-7edb-fc5f-3a2a7326bece
---
# Inspector.SetSchedulingStartTime メソッド (Outlook)(機械翻訳)




インスペクターの [**スケジュール アシスタント**] タブで、空き時間グリッドに会議アイテムの開始時刻を設定します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**SetSchedulingStartTime**(**Start**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Inspector** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Start*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**日付**|インスペクターの [**スケジュール アシスタント**] タブに会議の出席者の空き時間を表示する時間範囲の開始時刻を指定します。|





## 注釈
親の**Inspector**オブジェクトの**CurrentItem**プロパティによって指定されたオブジェクトは、 **AppointmentItem**または**MeetingItem**でなければなりません。インスペクターで、[**スケジュール アシスタント**] タブを表示する必要があります、それ以外の場合、 UNRESOLVED_TOKEN_VAL(outlooknv1)エラーが発生します。Outlook は、その項目の種類の [**スケジュール アシスタント**] タブを表示できない場合次のエラーが表示されます:**スケジュールの開始時刻は、スケジュール アシスタント会議アイテムに表示される場合にのみ設定できます**。



## 例
Microsoft Visual Basic for Applications (VBA) で次のコード サンプルでは、[**スケジュール アシスタント**] タブ、 **AppointmentItem**の開始時刻のスケジュールを設定するのには、 **SetSchedulingStartTime**メソッドを使用する方法を示します。予定の開始時刻は今から 1 か月に設定し、スケジュールの開始時間が今から 1 か月に設定されてもします。

```vba
Sub DemoSetSchedulingStartTime() 
 
 Dim oAppt As Outlook.AppointmentItem 
 
 Dim oInsp As Outlook.inspector 
 
 
 
 ' Create and display appointment. 
 
 Set oAppt = Application.CreateItem(olAppointmentItem) 
 
 oAppt.MeetingStatus = olMeeting 
 
 oAppt.Subject = "Test Appointment" 
 
 oAppt.Start = DateAdd("m", 1, Now) 
 
 ' Display the appointment in the Appointment tab of the inspector. 
 
 oAppt.Display 
 
 
 
 Set oInsp = oAppt.GetInspector 
 
 ' Switch to the Scheduling Assistant tab in that inspector. 
 
 oInsp.SetCurrentFormPage ("Scheduling Assistant") 
 
 ' Set the appointment start time in the Scheduling Assistant. 
 
 oInsp.SetSchedulingStartTime (DateAdd("m", 1, Now)) 
 
End Sub 
 

```




## Related Topics

[Inspector オブジェクトのメンバー](acd3e13f-4727-7966-d2a5-a95e4528425c.md)

[Inspector オブジェクト](d7384756-669c-0549-1032-c3b864187994.md)




