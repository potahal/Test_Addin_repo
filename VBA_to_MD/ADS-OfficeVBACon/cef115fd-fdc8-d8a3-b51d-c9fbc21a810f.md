

---
ms.Toctitle:CommandBarPopup.Priority プロパティ (Office)
title:CommandBarPopup.Priority プロパティ (Office)
ms.ContentId:cef115fd-fdc8-d8a3-b51d-c9fbc21a810f
---
# CommandBarPopup.Priority プロパティ (Office)




取得または**ポップアップ**コントロールの優先度を設定します。読み取り/書き込み。

## 

>[!NOTE]
>一部の Microsoft Office アプリケーションにおける CommandBars の使用方法が、Microsoft Office Fluent ユーザー インターフェイスの新しいリボン コンポーネントによって置き換えられました。詳細については、ヘルプでキーワード "リボン" を検索してください。





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Priority**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **CommandBarPopup** オブジェクトを表す変数を指定します。

### 戻り値
整数型 (Integer)





## 注釈
固定したコマンド バーのコントロールが 1 行に収まらない場合、コントロールの優先度に基づいて、コマンド バーの表示領域から削除されるコントロールが決定されます。1 行に収まらないコントロールは右のものから順番にコマンド バーから削除されます。



## 例
次の使用例は、コマンド バー ポップアップの説明文と優先度を設定します。

```sourcecode
Dim popControl As CommandBarPopup 
Set popControl = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics") 
 
With popControl. 
            .DescriptionText = "Graphics Selection dialog" 
            .Priority = 5 
End With 

```




>[!NOTE]
>
              UNRESOLVED_TOKEN_VAL(osdepreccommandbars)
            





## Related Topics

[ポップアップ オブジェクトのメンバー](8ec16deb-bb74-2871-d837-f706c7a58f2b.md)

[ポップアップ](a8ae06a3-1d7b-a531-91df-756fafee5314.md)




