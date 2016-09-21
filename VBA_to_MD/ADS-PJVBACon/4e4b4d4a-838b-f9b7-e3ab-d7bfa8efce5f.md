

---
ms.Toctitle:Application.SetAutoFilter メソッド (Project)
title:Application.SetAutoFilter メソッド (Project)
ms.ContentId:4e4b4d4a-838b-f9b7-e3ab-d7bfa8efce5f
---
# Application.SetAutoFilter メソッド (Project)




シート ビューで指定されたフィールドに対するオートフィルターの条件を設定します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**SetAutoFilter**(**FieldName**, **FilterType**, **Test1**, **Criteria1**, **Operation**, **Test2**, **Criteria2**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを返す式。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*FieldName*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|フィールドの名前を指定します。|
|*FilterType*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**PjAutoFilterType**|型のフィルターです。**PjAutoFilterType**定数のいずれかをすることができます。既定値は、 **pjAutoFilterClear**、オート フィルターをクリアします。|
|*Test1*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|最初のテストの比較の種類を指定します。*FilterType*は、 **pjAutoFilterCustom**、その*Criteria1*は、値を指定する必要があります。比較文字列は、以下のいずれかできます。比較文字列説明"と等しい" 引数 FieldName の値は、引数 Criteria1 の値と等しい。 "と等しくない" 引数 FieldName の値は、引数 Criteria1 の値と等しくない。 "より大きい" 引数 FieldName の値は、引数 Criteria1 の値より大きい。 "以上" 引数 FieldName の値は、引数 Criteria1 の値より大きいか等しい。 "より小さい" 引数 FieldName の値は、引数 Criteria1 の値より小さい。 "以下" 引数 FieldName の値は、引数 Criteria1 の値より小さいか等しい。 "の範囲内" 引数 FieldName の値は、引数 Criteria1 の値の範囲内にある。 "の範囲外" 引数 FieldName の値は、引数 Criteria1 の値の範囲内にない。|
|*Criteria1*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|FieldName で指定されるフィールドの値と比較する最初の比較の値を指定します。|
|*Operation*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|2 番目のテストがある場合に論理演算を指定します。Operation に指定できる値は、"かつ" または "または" です。|
|*Test2*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|2 番目のテストの比較の種類を指定します。*FilterType*は、 **pjAutoFilterCustom**、 *Operation*値を設定する必要があります、およびその*Criteria2*値を指定することが必要です。文字列には、Test1 のテーブルの比較のいずれかを指定できます。|
|*Criteria2*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|*FieldName* で指定されるフィールドの値と比較する 2 番目の比較の値を指定します。|



### 戻り値
**ブール型 (Boolean)**





## 注釈
オートフィルター機能をオンまたはオフにする方法については、**AutoFilter** メソッドを参照してください。

>[!NOTE]
>シート ビューの列名には、表示されるフィールドの名前とは別のタイトルを付けることができます。





## 例
次の使用例は、"作業時間の達成率" フィールドにユーザー定義のオートフィルターを設定します。

```vba
Sub TestAutoFilter() 
    If Not ActiveProject.AutoFilter Then 
        Application.AutoFilter 
    End If 
 
    Application.SetAutoFilter FieldName:="% Work Complete", FilterType:=pjAutoFilterCustom, _ 
    Test1:="equals", Criteria1:="0%" 
End Sub
```




ある場合は、「作業達成率] フィールドのオート フィルターの設定次のコード行は、オプション*FilterType*引数の既定値は**pjAutoFilterClear**であるため、オート フィルターをクリアします。

```vba
Application.SetAutoFilter FieldName:="% Work Complete"
```





