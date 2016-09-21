

---
ms.Toctitle:StartDriver.EffectiveDateAdd プロパティ (Project)
title:StartDriver.EffectiveDateAdd プロパティ (Project)
ms.ContentId:5b2e2c6e-06b9-ebf4-efdb-4ca2e944b7ff
---
# StartDriver.EffectiveDateAdd プロパティ (Project)




効果的なカレンダーを使用して手動でスケジュールされたタスクの指定された期間、別の日付に依存するときの日時を取得します。読み取り専用**バリアント**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**EffectiveDateAdd**(**Date**, **Duration**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **StartDriver** オブジェクトを返すオブジェクト式を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Date*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**バリアント型 (Variant)**|任意の日付と時刻 ("7/10/2010"、"7/10/2010 2:00:00 PM" など) を指定します。|
|*Duration*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**バリアント型 (Variant)**|追加する期間 ("3d"、"2w" など) を指定します。|





## 注釈
**EffectiveDateAdd**プロパティは、有効なカレンダーを手動でスケジュールされたタスクは、タスクが非稼働時間に開始および終了できるようにします。プロパティと引数なしに影響を与えるタスクの実際の日付。



開始日と計算終了日を手動でスケジュールされたタスクには、 **EffectiveDateSubtract**、 **EffectiveDateAdd**、および**EffectiveDateDifference**プロパティを使用できます。



カレンダーも指定できる自動でスケジュールされたタスクの日付を計算するには、**DateAdd** メソッドを使用してください。



## 例
次のステートメントは、指定した日付の 6 日後にあたる値 "7/9/2009 5:00:00 PM" を返します。

```vba
Debug.Print ActiveProject.Tasks(3).StartDriver.EffectiveDateAdd("7/2/2009", "6d")
```





