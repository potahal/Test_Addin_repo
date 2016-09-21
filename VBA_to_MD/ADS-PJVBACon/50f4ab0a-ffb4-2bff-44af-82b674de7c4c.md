

---
ms.Toctitle:Application.FilePageSetupCalendar メソッド (Project)
title:Application.FilePageSetupCalendar メソッド (Project)
ms.ContentId:50f4ab0a-ffb4-2bff-44af-82b674de7c4c
---
# Application.FilePageSetupCalendar メソッド (Project)




印刷するカレンダーのページを設定します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**FilePageSetupCalendar**(**Name**, **MonthsPerPage**, **WeeksPerPage**, **ScreenWeekHeight**, **OnlyDaysInMonth**, **OnlyWeeksInMonth**, **MonthPreviews**, **MonthTitle**, **AdditionalTasks**, **GroupAdditionalTasks**, **PrintNotes**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Name*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|印刷するカレンダーのページを設定します。|
|*MonthsPerPage*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|各ページに印刷する月の数です。1 または 2 を使用できます。**MonthsPerPage**引数は、 **OnlyDaysInMonth**または**OnlyWeeksInMonth**が指定されている場合に必要です。|
|*WeeksPerPage*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|1 ページに印刷する週数を指定します。|
|*ScreenWeekHeight*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True**印刷イメージを画面に表示される週の高さを使用する場合。|
|*OnlyDaysInMonth*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True**だけ、月の日付が印刷されます。**False**と、前の月の最後と次の月の開始日が現在の月の日付だけでなく印刷されます。**MonthsPerPage**の値が指定されていない限り、 **OnlyDaysInMonth**引数は無視されます。|
|*OnlyWeeksInMonth*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True**だけが月に完全に含まれる週を印刷します。**False**場合は、月に 1 つまたは複数の日のある週を印刷します。**MonthsPerPage**の値が指定されていない限り、 **OnlyWeeksInMonth**引数は無視されます。|
|*MonthPreviews*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True の**場合、前月および翌月のカレンダーを印刷します。|
|*MonthTitle*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**該当**月のタイトルを印刷する場合です。|
|*AdditionalTasks*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True**と予定表に表示されないタスクも印刷します。(追加のタスクは、印刷出力の末尾に表示)。|
|*GroupAdditionalTasks*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True**と、追加のタスクが 1 日でグループ化します。|
|*PrintNotes*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True の**各タスクに関連付けられているメモを印刷する場合です。ノートの他のタスクの後、最後に印刷します。|



### 戻り値
**ブール型 (Boolean)**





## 注釈
**FilePageSetupCalendar**メソッドを使用して引数を指定せず、[**表示**] タブで、[**ページ設定**] ダイアログ ボックスが表示されます。

    **FilePageSetupCalendar**メソッドを使用可能なは、カレンダーがアクティブなビューであるときのみです。



## 例
次の使用例は、1 ページに 2 か月のカレンダーを印刷し、前月と翌月のカレンダーをプレビューするように、カレンダーのページを設定します。

```vba
Sub File_PageSetupCalendar() 
 
 'Activate Calandar view 
 ViewApply Name:="&Calendar" 
 FilePageSetupCalendar MonthsPerPage:=2, OnlyDaysInMonth:=False, MonthPreviews:=True 
End Sub
```





