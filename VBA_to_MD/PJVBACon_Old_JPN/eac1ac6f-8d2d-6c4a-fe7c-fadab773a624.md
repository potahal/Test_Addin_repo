
# Application.DisplayPlanningWizard プロパティ (Project)

 **True** を指定すると、プランニング ウィザードがアクティブになります。値の取得および設定が可能です。ブール型 ( **Boolean** ) の値を使用します。


## 構文

 _式_. **DisplayPlanningWizard**

 _式_ **Application** オブジェクトを表す変数です。


## 例

次の使用例は、プランニング ウィザードの設定を既定値に戻します。


```
Sub ResetWizard() 
 Application.DisplayPlanningWizard = True 
 Application.DisplayWizardErrors = True 
 Application.DisplayWizardScheduling = True 
 Application.DisplayWizardUsage = True 
End Sub
```

