

---
ms.Toctitle:ComboBox.SpecialEffect プロパティ (Outlook フォーム スクリプト)
title:ComboBox.SpecialEffect プロパティ (Outlook フォーム スクリプト)
ms.ContentId:ac8f36a3-fd03-fa1a-4df1-ccc31bd059c9
---
# ComboBox.SpecialEffect プロパティ (Outlook フォーム スクリプト)




オブジェクトの外観を指定する**整数値**を設定または返します。 読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**SpecialEffect**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **ComboBox** オブジェクトを表す変数です。



## 注釈
**SpecialEffect**の設定は次のとおりです。

|**値**|**説明**|
|---|---|
|0|オブジェクトは平面的に表示され、境界線や色の違いによって背景と区別されます。|
|1|オブジェクトの上辺と左辺が強調表示になり、下辺と右辺には影が付けられます。|
|2|オブジェクトの上辺と左辺には影が付けられ、下辺と右辺は強調表示されます。コントロールと境界線は、それを包む曲線で表示されます。|
|3|コントロールの枠が沈んで見えます。|
|6|オブジェクトの下辺と右辺に隆起線が付けられ、上辺と左辺は平面的に表示されます。|



**SpecialEffect** ] または [**境界線スタイル**] プロパティを使用するには、コントロールが、両方の edging を指定します。これらのプロパティのいずれかの 0 以外の値を指定すると、0 に、他のプロパティの値が設定されます。**たとえば、 SpecialEffectを 1 に設定するとシステム設定を 0 に。**



**SpecialEffect**では、境界を定義するのにはシステム カラーを使用します。




