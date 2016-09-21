

---
ms.Toctitle:Image.BorderStyle プロパティ (Outlook フォーム スクリプト)
title:Image.BorderStyle プロパティ (Outlook フォーム スクリプト)
ms.ContentId:c3b9fb9e-76bb-3ad1-b75a-6acc03b4af9e
---
# Image.BorderStyle プロパティ (Outlook フォーム スクリプト)




コントロールの境界線の種類を指定する**整数値**を設定または返します。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**BorderStyle**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Image** オブジェクトを表す変数です。



## 注釈
**境界線スタイル**で使用できる値は、0 と 1 です。0 は境界線を表示がないことを表します、1 は、一重線の境界線 (既定値) を表します。



イメージ (**Image**) コントロールの既定値は 1 (1 枚) です。



コントロールが、両方の境界を指定するのには、 **BorderStyle**または**SpecialEffect**のいずれかを使用できます。これらのプロパティのいずれかの 0 以外の値を指定すると、0 に、他のプロパティの値が設定されます。たとえば、**境界線スタイル**を 1 に設定する場合システムは**SpecialEffect**をゼロ (フラット) に設定します。**SpecialEffect**の 0 以外の値を指定すると、システムはゼロに、**境界線スタイル**を設定します。



**境界線スタイル**は、その罫線の色を定義するのには**境界線色**を使用します。**BorderColor**プロパティを使用するには、 **BorderStyle**プロパティを 0 以外の値に設定しなければなりません。




