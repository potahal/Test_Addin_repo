

---
ms.Toctitle:以前の会議出席依頼への対抗案としての会議アイテムを識別します。
title:以前の会議出席依頼への対抗案としての会議アイテムを識別します。
ms.ContentId:42d53f48-d9de-18d8-d39b-86feceff0eaa
---
# 以前の会議出席依頼への対抗案としての会議アイテムを識別します。




このトピックでは、以前の会議出席依頼へのカウンターの提案として、 **MeetingItem**オブジェクトを識別する名前付きのプロパティ、 [PidLidAppointmentCounterProposal](f510af2d-92b3-4c98-bdf4-8aca8e8b80d1.md)、および Microsoft Outlook オブジェクト モデルを使用する方法を示します。



Outlook オブジェクト モデルでは、すべての種類のアイテムがメール アイテムや連絡先アイテムなどの特定のメッセージ クラスと対応しています。具体的には、会議出席依頼への応答は次のメッセージ クラスで識別できます。 

- 辞退の応答は IPM.Schedule.Meeting.Resp.Neg
- 出席の応答は IPM.Schedule.Meeting.Resp.Pos
- 仮承諾の応答は IPM.Schedule.Meeting.Resp.Ten








ただし、Outlook オブジェクト モデルにはあり得る 4 つ目の応答を識別する手段がありません。それは、新しい日次の指定です。



**PropertyAccessor**オブジェクトと**PidLidAppointmentCounterProposal**の**PSETID_Appointment**の名前空間の定義を使用すると、プログラミングできるオブジェクト モデル内での会議のすべての応答を区別するためにアイテムを要求します。C# で次のコード サンプルでは、会議アイテムを指定されたプロパティ値を取得する方法を示します。コード サンプルでは、名前付きプロパティとして表されることに注意してください。

```sourcecode
"http://schemas.microsoft.com/mapi/id/00062002-0000-0000-C000-000000000046}/8257000B"
```




`{00062002-0000-0000-C000-000000000046}` は **PSETID_Appointment** 名前空間で、`8257000B` は **PidLidAppointmentCounterProposal** のプロパティ タグです。

```csharp
private bool IsCounterProposal(Outlook.MeetingItem meeting) 
{ 
    const string counterPropose = 
        "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8257000B"; 
    Outlook.PropertyAccessor pa = meeting.PropertyAccessor; 
    if ((bool)pa.GetProperty(counterPropose)) 
        return true; 
    else 
        return false;  
}
```



