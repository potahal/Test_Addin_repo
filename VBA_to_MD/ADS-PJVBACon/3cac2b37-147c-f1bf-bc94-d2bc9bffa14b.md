

---
ms.Toctitle:Application.GanttBarStyleDelete メソッド (Project)
title:Application.GanttBarStyleDelete メソッド (Project)
ms.ContentId:3cac2b37-147c-f1bf-bc94-d2bc9bffa14b
---
# Application.GanttBarStyleDelete メソッド (Project)




作業中の [ガント チャート] ビューのガント バーのスタイルを削除します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GanttBarStyleDelete**(**Item**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Item*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|**文字列**です。**バーのスタイル**] ダイアログ ボックスから削除するガント バーの名前または行の数です。|



### 戻り値
**ブール型 (Boolean)**





## 注釈
[**バーのスタイル**] ダイアログ ボックスを手動で表示するには、[**ガント チャートのツール**] タブの下の [**形式**] タブをクリックします。[**バーのスタイル**] で、[**形式**] ボックスの一覧の [**バーのスタイル**] をクリックします。[**バーのスタイル**] ダイアログ ボックスには最大 200 のスタイルを登録できます。



## 例
次のコマンドは、[**バーのスタイル**] ダイアログ ボックスのスタイル番号 41 を削除します。

```vba
GanttBarStyleDelete Item:="41"
```





