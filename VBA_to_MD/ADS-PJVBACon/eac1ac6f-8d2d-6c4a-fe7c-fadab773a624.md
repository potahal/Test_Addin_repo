

---
ms.Toctitle:Application.DisplayPlanningWizard プロパティ (Project)
title:Application.DisplayPlanningWizard プロパティ (Project)
ms.ContentId:eac1ac6f-8d2d-6c4a-fe7c-fadab773a624
---
# Application.DisplayPlanningWizard プロパティ (Project)




**True** を指定すると、プランニング ウィザードがアクティブになります。値の取得および設定が可能です。ブール型 (**Boolean**) の値を使用します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**DisplayPlanningWizard**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。



## 例
次の使用例は、プランニング ウィザードの設定を既定値に戻します。

```vba
Sub ResetWizard() 
 Application.DisplayPlanningWizard = True 
 Application.DisplayWizardErrors = True 
 Application.DisplayWizardScheduling = True 
 Application.DisplayWizardUsage = True 
End Sub
```





