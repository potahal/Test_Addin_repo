

---
ms.Toctitle:Outlook アイテムを分類する
title:Outlook アイテムを分類する
ms.ContentId:e8cfb450-b8b0-bee6-fdf0-d0a92bf9af56
---
# Outlook アイテムを分類する





          UNRESOLVED_TOKEN_VAL(outlooknv1) では、Outlook アイテムを分類項目別に分類および表示できる色分類機能が用意されています。複数の色分類項目を 1 つの Outlook アイテムに適用できるほか、Outlook アイテムを色分類項目別にグループ化または並べ替えできます。各色分類項目にはショートカット キーを割り当てることができるため、アイテムをより簡単に分類できます。色分類項目はユーザー定義で、その作成、削除、および変更は、プログラム上でも、Outlook のユーザー インターフェイスにおけるユーザー操作でも実行できます。



**Category** オブジェクトは、分類項目マスター (Outlook のユーザー インターフェイスに表示され、**NameSpace** オブジェクトの **Categories** コレクションで表される色分類項目の一覧) に含まれる、1 つのユーザー定義の色分類項目を表します。**Category** オブジェクトは、作成時にグローバル識別子 (GUID) で識別されますが、この識別子は変更できません。しかし、色分類項目に関連付けられた名前、色、およびショートカット キーは、それぞれ **Category** オブジェクトの **Name**、**Color**、**ShortcutKey** の各プロパティを設定することで変更できます。**CategoryID** プロパティを使用すると、**Category** オブジェクトの識別子を取得できます。

## Outlook アイテムに分類項目を割り当てる
Outlook アイテムに分類項目を割り当てるには、次のオブジェクトの **Categories** プロパティに、該当する **Category** オブジェクトの名前を、コンマ区切り形式の文字列で指定します。

|||
|---|---|
|**AppointmentItem**|**RemoteItem**|
|**ContactItem**|**ReportItem**|
|**DistListItem**|**SharingItem**|
|**DocumentItem**|**PostItem**|
|**JournalItem**|**TaskItem**|
|**MailItem**|**TaskRequestAcceptItem**|
|**MeetingItem**|**TaskRequestDeclineItem**|
|**MobileItem**|**TaskRequestItem**|
|**NoteItem**|**TaskRequestUpdateItem**|



Outlook アイテムは、その Outlook アイテムの **Categories** プロパティに指定された分類項目名に基づいて表示されます。分類項目名は Outlook アイテムの一部として保存されるため、分類項目マスターに存在しない分類項目名を Outlook アイテムに指定することもできます。たとえば、分類項目が削除された場合などがこれに該当します。



対応する **Name** プロパティの値を持つ **Category** オブジェクトが、Outlook アイテムを含む **NameSpace** オブジェクトの **Categories** コレクションに存在しない場合、Outlook アイテムに関連付けられたその分類項目名は表示されますが、色は関連付けられません。




