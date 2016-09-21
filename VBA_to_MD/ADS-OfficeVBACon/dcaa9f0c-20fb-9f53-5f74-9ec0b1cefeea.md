

---
ms.Toctitle:COMAddIn オブジェクト (Office)
title:COMAddIn オブジェクト (Office)
ms.ContentId:dcaa9f0c-20fb-9f53-5f74-9ec0b1cefeea
---
# COMAddIn オブジェクト (Office)




Microsoft Office のホスト アプリケーションの COM アドインを表します。**COMAddIn**オブジェクトは、**COMAddIns** コレクションのメンバーです。

## 次の使用例では、テーブルからレコードを削除できないようにします。
**COMAddIns.Item(index)** を使用します。この場合の *index* は、**COMAddIns** コレクションの位置で COM アドインを返す序数値か、指定された COM アドインの ProgID を表す **String** 値です。次の例では、COM アドインを説明するテキストをメッセージ ボックスに表示します。

```sourcecode
MsgBox Application.COMAddIns.Item("msodraa9.ShapeSelect").Description
```




COM アドインのプログラム識別子を取得するには、 **COMAddin** オブジェクトの **ProgID** プロパティを使用し、COM アドインのグローバル一意識別子 (GUID) を取得するには、**Guid** プロパティを使用します。次の例では、COM アドイン 1 の ProgID と GUID をメッセージ ボックスに表示します。

```sourcecode
MsgBox "My ProgID is " & _ 
 Application.COMAddIns(1).ProgID & _ 
 " and my GUID is " & _ 
 Application.COMAddIns(1).Guid
```




指定した COM アドインに対する接続の状態を設定または取得するには、**Connect** プロパティを使用します。次の例では、COM アドイン 1 が登録され、現在接続されているかどうかを示すメッセージ ボックスを表示します。

```sourcecode
If Application.COMAddIns(1).Connect Then 
 MsgBox "The add-in is connected." 
Else 
MsgBox "The add-in is not connected." 
End If
```




## Related Topics

[Object Model Reference](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)

[COMAddIn Object Members](698d4d8e-6071-acd3-a39b-ab01fd878452.md)




