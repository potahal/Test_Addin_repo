

---
ms.Toctitle:Application.OptionsInterfaceEx メソッド (Project)
title:Application.OptionsInterfaceEx メソッド (Project)
ms.ContentId:da4dc69c-021f-7ecb-22f6-aebf1d9252dd
---
# Application.OptionsInterfaceEx メソッド (Project)




表示オプションおよびプロジェクト ガイドのオプションをいくつか設定します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**OptionsInterfaceEx**(**ShowResourceAssignmentIndicators**、**ShowEditToStartFinishDates**、**ShowEditsToWorkUnitsDurationIndicators**、**ShowDeletionInNameColumn**、**DisplayProjectGuide**、**ProjectGuideUseDefaultFunctionalLayoutPage**、**ProjectGuideFunctionalLayoutPage**、**ProjectGuideUseDefaultContent**、**ProjectGuideContent**、**SetAsDefaults**、**UseOMIDs**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*ShowResourceAssignmentIndicators*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True** を指定すると、リソース割り当てのマークとオプション ボタンが表示されます。既定値は **False** です。|
|*ShowEditToStartFinishDates*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True** を指定すると、開始日と終了日の編集を元に戻すスタックにアクションが表示されます。既定値は **False** です。|
|*ShowEditsToWorkUnitsDurationIndicators*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True** を指定すると、期間の編集を元に戻すスタックにアクションが表示されます。既定値は **False** です。|
|*ShowDeletionInNameColumn*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True** を指定すると、Project の [**タスク名**] または [**リソース名**] フィールドの値を削除した後、元に戻すスタックにアクションが表示されます。既定値は **False** です。|
|*DisplayProjectGuide*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|
						起動時およびすべての新規プロジェクトに既定で **Project Guide** を表示する必要がある場合、**True** を指定します。既定値は False です。|
|*ProjectGuideUseDefaultFunctionalLayoutPage*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True** を指定すると、プロジェクト ガイドでは既定のコンテンツが使用されます。**False** を指定すると、プロジェクト ガイド用のカスタム コンテンツが使用されます。既定値は **True** です。|
|*ProjectGuideFunctionalLayoutPage*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|**Project Guide** の値を指定します。独自のコンテンツで使用する XML ファイルの URL またはパスと名前を指定します。|
|*ProjectGuideUseDefaultContent*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True** を指定すると、**プロジェクト ガイド** は既定のコンテンツを使用します。**False** を指定すると、プロジェクト ガイドはカスタム コンテンツを使用します。既定値は **True** です。|
|*ProjectGuideContent*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|プロジェクト ガイドの値を指定します。カスタム コンテンツで使用する XML ファイルの URL またはパスと名前を指定します。|
|*SetAsDefaults*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True** を指定すると、作業中のプロジェクトの **Project Guide** の設定が、新しいすべてのプロジェクトの既定値として使用されます。既定値は False です。|
|*UseOMIDs*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|**True** を指定すると、プロジェクト間で言語や名前が異なる構成アイテムを一致させるため、Project で内部 ID が使用されます。既定値は **True** です。**UseOMIDs** プロパティも参照してください。|



### 戻り値
**Boolean**





## 注釈
引数を省略すると、既定値は [**Project のオプション**] ダイアログ ボックスの [**表示**] タブの設定で指定されます。*UseOMIDs* の既定値は [**詳細**] タブの [**内部 ID を使用する**] オプションです。

>[!NOTE]
>**プロジェクトのオプション**] ダイアログ ボックスでは、 UNRESOLVED_TOKEN_VAL(pjgenericshort)では使用されなくなりましたが、プロジェクト ガイドの設定は含まれません。プロジェクト ガイドのオプションは、カスタムのプロジェクト ガイドを使用するプログラムでのみ設定できます。新しいプロジェクト ガイドのコンテンツを作成するのではなく開発者タスクが作成されますアプリケーションのウィンドウです。





引数を指定しないで **OptionsInterfaceEx** メソッドを使用すると、[**一般**] タブが選択された状態で [**プロジェクトのオプション**] ダイアログボックスが表示されます。レポート ビューでの作業中は、**OptionsInterfaceEx** メソッドは使用できません。




