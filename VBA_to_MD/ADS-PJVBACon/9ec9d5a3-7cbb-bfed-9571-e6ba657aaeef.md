

---
ms.Toctitle:Application.GanttBarFormatEx メソッド (Project)
title:Application.GanttBarFormatEx メソッド (Project)
ms.ContentId:9ec9d5a3-7cbb-bfed-9571-e6ba657aaeef
---
# Application.GanttBarFormatEx メソッド (Project)




ガント バーの書式を既定のスタイルから変更します。バーの色は、16 進数の RGB 値で指定できます。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GanttBarFormatEx**(**TaskID**, **GanttStyle**, **StartShape**, **StartType**, **StartColor**, **MiddleShape**, **MiddlePattern**, **MiddleColor**, **EndShape**, **EndType**, **EndColor**, **LeftText**, **RightText**, **TopText**, **BottomText**, **InsideText**, **Reset**, **ProjectName**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを返す式。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*TaskID*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**長整数型 (Long)**|ガント チャートのスタイルを変更するタスクの ID 番号を指定します。既定では、選択したタスクのガント バーが変更されます。|
|*GanttStyle*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|書式設定するガント バーに適用されるスタイルです。GanttStyle の値は、バーの位置に基づいて、[**バーのスタイル**] ダイアログ ボックスのスタイル。たとえば、値 3 を返します 3 番目のバー スタイルの一覧にします。|
|*StartShape*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|ガント バーの左端の形状を指定します。使用できる定数は、**PjBarEndShape** クラスの定数のいずれかです。|
|*StartType*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|ガント バーの左端の種類を指定します。使用できる定数は、**PjBarType** クラスの定数のいずれかです。|
|*StartColor*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**長整数型 (Long)**|ガント バーの左端の形状の色を指定します。RGB 色を 16 進数の値で指定し、最後のバイトが赤色を表します。たとえば、値 &H00FFFF は黄色を表します。|
|*MiddleShape*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|ガント バーのバーの形状を指定します。使用できる定数は、**PjBarShape** クラスの定数のいずれかです。|
|*MiddlePattern*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|ガント バーのバーのパターンを指定します。使用できる定数は、**PjFillPattern** クラスの定数のいずれかです。|
|*MiddleColor*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**長整数型 (Long)**|ガント バーの中央部分の色を指定します。RGB 色を 16 進数の値で指定し、最後のバイトが赤色を表します。たとえば、値 &HFF00FF は紫色を表します。|
|*EndShape*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|ガント バーの右端の形状です。**PjBarEndShape**定数のいずれかをすることができます。|
|*EndType*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|ガント バーの右端の種類です。**PjBarType**定数は、次のいずれか: **pjDashed**、 **pjFramed**、または**pjSolid**です。|
|*EndColor*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**長整数型 (Long)**|ガント バーの右端の形状の色を指定します。RGB 色を 16 進数の値で指定し、最後のバイトが赤色を表します。たとえば、値 &HFFFF00 は青緑色を表します。|
|*LeftText*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|ガント バーの左側に表示するタスク フィールドの名前を指定します。|
|*RightText*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|ガント バーの右側に表示するタスク フィールドの名前を指定します。|
|*TopText*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|ガント バーの上側に表示するタスク フィールドの名前を指定します。|
|*BottomText*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|ガント バーの下側に表示するタスク フィールドの名前を指定します。|
|*InsideText*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|ガント バーの内側に表示するタスク フィールドの名前を指定します。|
|*Reset*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True**の場合、バーの形式は、 **[バーのスタイル**ダイアログ ボックスの [スタイルの既定の書式設定にリセットがそれ以外の場合、 **false を指定**します。|
|*ProjectName*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|統合されている場合、**引数 TaskID**を含むプロジェクトの名前です。既定値は、作業中のプロジェクトの名前です。|



### 戻り値
**ブール型 (Boolean)**





## 注釈
引数を指定せず、 **GanttBarFormatEx**メソッドを使用するには、**バーの書式設定**] ダイアログ ボックスが表示されます。



色を 16 進数の RGB 値で指定できる既定のスタイルを定義するには、**GanttBarEditEx** メソッドを使用します。



## 例
次の使用例は、タスク ID が 3 のタスクの開始点にやや濃い赤色のひし形を表示します。

```vba
Sub GanttBar_Format() 
 
    'Activate Gantt Chart view 
    ViewApply Name:="&Gantt Chart" 
    GanttBarFormatEx TaskID:=3, StartShape:=pjDiamond, StartType:=pjSolid, StartColor:=&H8888FF
End Sub
```





