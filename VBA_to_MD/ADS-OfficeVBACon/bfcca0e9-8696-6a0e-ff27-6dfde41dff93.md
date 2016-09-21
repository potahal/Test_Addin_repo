

---
ms.Toctitle:IRibbonUI.InvalidateControlMso メソッド (Office)
title:IRibbonUI.InvalidateControlMso メソッド (Office)
ms.ContentId:bfcca0e9-8696-6a0e-ff27-6dfde41dff93
---
# IRibbonUI.InvalidateControlMso メソッド (Office)




組み込みのコントロールを無効にするのに使用します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**InvalidateControlMso**(**ControlID**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **IRibbonUI** オブジェクトを表すオブジェクト式を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*ControlID*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**||



### 戻り値
なし





## 注釈
コントロールを無効にすると、画面が再描画され、そのコントロールに関連付けられているすべてのコールバック プロシージャが実行されます。



## 例

```xml
<customUI … OnLoad=”MyAddInInitialize” …>
```


```vba
Sub MyAddInInitialize(Ribbon As IRibbonUI) 
 Set MyRibbon = Ribbon 
End Sub 
 
Sub myFunction() 
 MyRibbon.InvalidateControlMso("TabInsert") ‘ Invalidates the Insert control 
End Sub
```




## Related Topics

[IRibbonUI オブジェクト](d323aa21-de74-e821-c914-db71ef3b9c5e.md)

[IRibbonUI オブジェクトのメンバー](c6f6ec3b-3132-da29-ea08-70f20923d013.md)




