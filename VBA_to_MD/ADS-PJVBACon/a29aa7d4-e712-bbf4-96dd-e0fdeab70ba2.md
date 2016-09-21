

---
ms.Toctitle:Pane.View メソッド (Project)
title:Pane.View メソッド (Project)
ms.ContentId:a29aa7d4-e712-bbf4-96dd-e0fdeab70ba2
---
# Pane.View メソッド (Project)




アクティブな**ビュー**オブジェクトを返します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**View**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Pane** オブジェクトを表す変数です。

### 戻り値
**表示**





## 例
次のステートメントは、VBE の[**イミディエイト**] ウィンドウにビューの名前を出力します。たとえば、[チーム プランナー] ビューがアクティブな場合は、"Team Plannner" と出力されます。

```vba
Debug.Print ActiveWindow.ActivePane.View.Name
```





