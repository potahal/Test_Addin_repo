

---
ms.Toctitle:モジュールをナビゲーション ウィンドウで現在選択されているモジュールとして設定します。
title:モジュールをナビゲーション ウィンドウで現在選択されているモジュールとして設定します。
ms.ContentId:c7aeafcf-d88d-8d79-8dfd-e336cf00f101
---
# モジュールをナビゲーション ウィンドウで現在選択されているモジュールとして設定します。




**NavigationPane** オブジェクトの **CurrentModule** プロパティを使用すると、UNRESOLVED_TOKEN_VAL(outlooknv1) で **Explorer** オブジェクトのナビゲーション ウィンドウで **NavigationModule** オブジェクトを現在選択されているナビゲーション モジュールとして設定できます。



次のサンプルは、[**履歴**] ナビゲーション モジュールが選択されている場合、ナビゲーション ウィンドウでプログラムまたはユーザーの操作によって [**予定表**] ナビゲーション モジュールを現在選択されているナビゲーション モジュールとして設定します。実行する処理は以下のとおりです。

1. **Application** オブジェクトの **Startup** イベントが発生したときに、アクティブなエクスプローラーの **NavigationPane** オブジェクトへの参照を取得し、それを `objPane` に代入し、**NavigationPane** オブジェクトの **ModuleSwitch** イベントを検出できるようにします。
2. **NavigationPane** の **ModuleSwitch** イベントが発生すると、**ModuleSwitch** イベントの *CurrentModule* パラメーターの内容と **NavigationPane** オブジェクトの **CurrentModule** プロパティを比較して、現在のナビゲーション モジュールが変更されているかどうかを確認します。
3. これらのオブジェクト参照が異なる場合、**ModuleSwitch** イベントの *CurrentModule* パラメーターで **NavigationModule** オブジェクト参照の **NavigationModuleType** プロパティを確認します。
4. 現在選択されている **Module** オブジェクトの **NavigationModuleType** プロパティが **olModuleJournal** に設定されている場合、現在選択されている [**履歴**] ナビゲーション モジュールが一時的に使用不可であり、代わりに [**予定表**] ナビゲーション モジュールが選択されることをユーザーに示すダイアログ ボックスを表示します。
5. 最後に、**NavigationPane** オブジェクトの **Modules** コレクションの **GetNavigationModule** メソッドを使用して、**CalendarModule** オブジェクトの取得を試みます。成功した場合、**NavigationPane** オブジェクトの **CurrentModule** プロパティは、取得した **CalendarModule** オブジェクト参照に設定されます。


```sourcecode
Dim WithEvents objPane As NavigationPane 
 
Private Sub Application_Startup() 
 ' Get the NavigationPane object for the 
 ' currently displayed Explorer object. 
 Set objPane = Application.ActiveExplorer.NavigationPane 
 
End Sub 
 
Private Sub objPane_ModuleSwitch(ByVal CurrentModule As NavigationModule) 
 Dim objModule As CalendarModule 
 
 ' Check if the currently selected navigation module 
 ' has changed. 
 If Not (CurrentModule Is objPane.CurrentModule) Then 
 ' If the Journal module was selected, forcibly change 
 ' it to the Calendar module by setting the 
 ' CurrentModule property of the NavigationPane object. 
 If CurrentModule.NavigationModuleType = olModuleJournal Then 
 
 ' Let the user know what's happening. 
 MsgBox "The Journal module is temporarily unavailable. " & _ 
 " Outlook is switching to the Calendar module, if available." 
 
 ' Retrieve the Calendar module, if one exists, for the 
 ' current Navigation Pane. 
 Set objModule = objPane.Modules.GetNavigationModule(olModuleCalendar) 
 
 ' If we have one, set the CurrentModule property of the 
 ' NavigationPane object to the Calendar module. 
 If Not (objModule Is Nothing) Then 
 Set objPane.CurrentModule = objModule 
 End If 
 End If 
 End If 
 
End Sub 

```



