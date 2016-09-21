

---
ms.Toctitle:マルチ ページ (MultiPage) コントロールとタブ ストリップ (TabStrip) コントロールの、タブのスタイルを設定する方法
title:マルチ ページ (MultiPage) コントロールとタブ ストリップ (TabStrip) コントロールの、タブのスタイルを設定する方法
ms.ContentId:abaeceb2-d8fd-92a5-8d54-eab94d9de394
---
# マルチ ページ (MultiPage) コントロールとタブ ストリップ (TabStrip) コントロールの、タブのスタイルを設定する方法




次の例は、**Style** プロパティを使用して、マルチ ページ (**MultiPage**) コントロールとタブ ストリップ (**TabStrip**) コントロールのタブの外観を指定します。この例では、ラベル (**Label**) コントロールも使用して示します。オプション ボタン (**OptionButton**) コントロールを選択して、スタイルを選びます。



この例を利用するには、次のコード例をフォームのスクリプト エディターにコピーします。コードを実行するには、**Open** イベントが生じるようにフォームを開く必要があります。フォームには次のコントロールが含まれている必要があります。

- ラベル (**Label**) コントロール (Label1)
- 3 つのオプション ボタン (**OptionButton**) コントロール (OptionButton1、OptionButton2、OptionButton3)
- マルチ ページ (**MultiPage**) コントロール (MultiPage1)
- タブ ストリップ (**TabStrip**) コントロール (TabStrip1)
- タブ ストリップ (**TabStrip**) コントロール内に任意のコントロール
- マルチ ページ (**MultiPage**) コントロールの各ページに任意のコントロール


```sourcecode
Sub OptionButton1_Click() 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 MultiPage1.Style = 0 '0=fmTabStyleTabs 
 TabStrip1.Style = 0 '0=fmTabStyleTabs 
End Sub 
 
Sub OptionButton2_Click() 
 'Note that the page borders are invisible 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 MultiPage1.Style = 1 '1=fmTabStyleButtons 
 TabStrip1.Style = 1 '1=fmTabStyleButtons 
End Sub 
 
Sub OptionButton3_Click() 
 'Note that the page borders are invisible and 
 'the page body begins where the tabs normally appear. 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 MultiPage1.Style = 2 '2=fmTabStyleNone 
 TabStrip1.Style = 2 '2=fmTabStyleNone 
End Sub 
 
Sub Item_Open() 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label1") 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton1") 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton2") 
 Set OptionButton3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton3") 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 
 Label1.Caption = "Page/Tab Style" 
 OptionButton1.Caption = "Tabs" 
 OptionButton1.Value = True 
 MultiPage1.Style = 0 '0=fmTabStyleTabs 
 TabStrip1.Style = 0 '0=fmTabStyleTabs 
 
 OptionButton2.Caption = "Buttons" 
 OptionButton3.Caption = "No Tabs or Buttons" 
End Sub
```



