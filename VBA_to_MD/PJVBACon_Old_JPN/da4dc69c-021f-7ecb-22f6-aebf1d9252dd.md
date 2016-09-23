
# Application.OptionsInterfaceEx メソッド (Project)

表示オプションおよびプロジェクト ガイドのオプションをいくつか設定します。


## 構文

 _式_. **OptionsInterfaceEx**( ** _ShowResourceAssignmentIndicators_** 、 ** _ShowEditToStartFinishDates_** 、 ** _ShowEditsToWorkUnitsDurationIndicators_** 、 ** _ShowDeletionInNameColumn_** 、 ** _DisplayProjectGuide_** 、 ** _ProjectGuideUseDefaultFunctionalLayoutPage_** 、 ** _ProjectGuideFunctionalLayoutPage_** 、 ** _ProjectGuideUseDefaultContent_** 、 ** _ProjectGuideContent_** 、 ** _SetAsDefaults_** 、 ** _UseOMIDs_** )

 _式_ **Application** オブジェクトを表す変数です。


### パラメーター



|**名前**|**必須/オプション**|**データ型**|**説明**|
|:-----|:-----|:-----|:-----|
| _ShowResourceAssignmentIndicators_|省略可能|**ブール型 (Boolean)**|**True** を指定すると、リソース割り当てのマークとオプション ボタンが表示されます。既定値は **False** です。|
| _ShowEditToStartFinishDates_|省略可能|**ブール型 (Boolean)**|**True** を指定すると、開始日と終了日の編集を元に戻すスタックにアクションが表示されます。既定値は **False** です。|
| _ShowEditsToWorkUnitsDurationIndicators_|省略可能|**ブール型 (Boolean)**|**True** を指定すると、期間の編集を元に戻すスタックにアクションが表示されます。既定値は **False** です。|
| _ShowDeletionInNameColumn_|省略可能|**ブール型 (Boolean)**|**True** を指定すると、Project の [ **タスク名**] または [ **リソース名**] フィールドの値を削除した後、元に戻すスタックにアクションが表示されます。既定値は  **False** です。|
| _DisplayProjectGuide_|省略可能|**ブール型 (Boolean)**|起動時およびすべての新規プロジェクトに既定で  **Project Guide** を表示する必要がある場合、 **True** を指定します。既定値は False です。|
| _ProjectGuideUseDefaultFunctionalLayoutPage_|省略可能|**ブール型 (Boolean)**|**True** を指定すると、プロジェクト ガイドでは既定のコンテンツが使用されます。 **False** を指定すると、プロジェクト ガイド用のカスタム コンテンツが使用されます。既定値は **True** です。|
| _ProjectGuideFunctionalLayoutPage_|省略可能|**文字列型 (String)**|**Project Guide** の値を指定します。独自のコンテンツで使用する XML ファイルの URL またはパスと名前を指定します。|
| _ProjectGuideUseDefaultContent_|省略可能|**ブール型 (Boolean)**|**True** を指定すると、 **プロジェクト ガイド** は既定のコンテンツを使用します。 **False** を指定すると、プロジェクト ガイドはカスタム コンテンツを使用します。既定値は **True** です。|
| _ProjectGuideContent_|省略可能|**文字列型 (String)**|プロジェクト ガイドの値を指定します。カスタム コンテンツで使用する XML ファイルの URL またはパスと名前を指定します。|
| _SetAsDefaults_|省略可能|**ブール型 (Boolean)**|**True** を指定すると、作業中のプロジェクトの **Project Guide** の設定が、新しいすべてのプロジェクトの既定値として使用されます。既定値は False です。|
| _UseOMIDs_|省略可能|**バリアント型 (Variant)**|**True** を指定すると、プロジェクト間で言語や名前が異なる構成アイテムを一致させるため、Project で内部 ID が使用されます。既定値は **True** です。 **[UseOMIDs](15339e09-0b65-d939-df47-eb538dee7c38.md)** プロパティも参照してください。|

### 戻り値

 **Boolean**


## 注釈

引数を省略すると、既定値は [ **Project のオプション**] ダイアログ ボックスの [ **表示**] タブの設定で指定されます。 _UseOMIDs_ の既定値は [ **詳細**] タブの [ **内部 ID を使用する**] オプションです。


 **メモ**   **プロジェクトのオプション**] ダイアログ ボックスでは、 Projectでは使用されなくなりましたが、プロジェクト ガイドの設定は含まれません。プロジェクト ガイドのオプションは、カスタムのプロジェクト ガイドを使用するプログラムでのみ設定できます。新しいプロジェクト ガイドのコンテンツを作成するのではなく開発者タスクが作成されますアプリケーションのウィンドウです。

引数を指定しないで  **OptionsInterfaceEx** メソッドを使用すると、[ **一般**] タブが選択された状態で [ **プロジェクトのオプション**] ダイアログボックスが表示されます。レポート ビューでの作業中は、 **OptionsInterfaceEx** メソッドは使用できません。

