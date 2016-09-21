

---
ms.Toctitle:CommandBarComboBox オブジェクト (Office)
title:CommandBarComboBox オブジェクト (Office)
ms.ContentId:fcfe6bde-dea0-f1f1-ad30-d0e28f97dd07
---
# CommandBarComboBox オブジェクト (Office)




コマンド バーのコンボ ボックス コントロールを表します。

## 

>[!NOTE]
>一部の Microsoft Office アプリケーションにおける CommandBars の使用方法が、Microsoft Office Fluent ユーザー インターフェイスの新しいリボン コンポーネントによって置き換えられました。詳細については、ヘルプでキーワード "リボン" を検索してください。





## 注釈
**戻します**オブジェクトを取得するのにには、 **Controls(index)**、*インデックス*が、コントロールのインデックス番号を使用します。******MsoControlDropdown**、 **msoControlComboBox**、 **msoControlButtonDropdown**、 **msoControlSplitDropdown**、 **msoControlOCXDropdown**、 **msoControlGraphicCombo**、または**インデックス番号を表します**に、コントロールの**Type**プロパティがある必要があることに注意してください。



## 例
次の使用例は、"Custom" というコマンド バーの 2 番目のコントロールに 2 つの項目を追加し、そのコントロールのサイズを調整します。

```sourcecode
Set combo = CommandBars("Custom").Controls(2) 
With combo 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListIndex = 0 
End With
```




**戻します**オブジェクトを取得するのに**FindControl**メソッドを使用することもできます。次の例が表示されている**戻します**オブジェクトのタグは、「シートの割り当て」のすべてのコマンド バーを検索します。

```sourcecode
Set myControl = CommandBars.FindControl _ 
(Type:=msoControlComboBox, Tag:="sheet assignments", Visible:=True)
```




>[!NOTE]
>
              UNRESOLVED_TOKEN_VAL(osdepreccommandbars)
            





## Related Topics

[オブジェクト モデル リファレンス](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)

[戻しますオブジェクトのメンバー](223c51c0-4564-d14a-a8bf-d315a6a50b32.md)




