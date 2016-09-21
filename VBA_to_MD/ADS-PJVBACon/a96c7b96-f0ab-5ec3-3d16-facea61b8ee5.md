

---
ms.Toctitle:Calendars オブジェクト (Project)
title:Calendars オブジェクト (Project)
ms.ContentId:a96c7b96-f0ab-5ec3-3d16-facea61b8ee5
---
# Calendars オブジェクト (Project)




**Calendar** オブジェクトのコレクションを格納します。

## 例
**Calendar オブジェクトの使い方**



**Calendar** オブジェクトを取得するには、**BaseCalendars(***Index***)** を使用します。引数 *Index* にはカレンダーのインデックス番号またはカレンダー名を指定します。

```vba
MsgBox ActiveProject.BaseCalendars(1).Name
```




**Calendars コレクションの使い方**



[Calendars](fb7f55f6-6618-fb82-dae1-320953bcf79d.md) コレクションを取得するには、**BaseCalendars** プロパティを使用します。次の使用例は、作業中のプロジェクトの各基本カレンダーのプロパティを既定値に戻します。

```vba
Dim C As Calendar 

 

For Each C In ActiveProject.BaseCalendars 

 C.Reset 

Next C
```




[Calendar](c9c92dff-255a-041b-c18d-49d6d75884e3.md) オブジェクトを **Calendars** コレクションに追加するには、**BaseCalendarCreate** メソッドを使用します。次の使用例は、新しい基本カレンダーを作成します。

```vba
BaseCalendarCreate Name:="Base Holiday Calendar"
```




## Related Topics

[プロジェクト オブジェクト モデル](900b167b-88ec-ea88-15b7-27bb90c22ac6.md)




