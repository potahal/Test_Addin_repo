

---
ms.Toctitle:GradientStops.Insert2 メソッド (Office)
title:GradientStops.Insert2 メソッド (Office)
ms.ContentId:bd9ed41d-eaeb-d3aa-6a8a-e38e2bfb9a17
---
# GradientStops.Insert2 メソッド (Office)




グラデーションに分岐点を追加し、分岐点の色の明るさおよび透明度を指定します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Insert2**(**RGB**, **Position**, **Transparency**, **Index**, **Brightness**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **GradientStops** オブジェクトを返す式を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*RGB*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**MsoRGBType**|グラデーションの分岐点の色を指定します。|
|*Position*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**単精度浮動小数点型 (Single)**|グラデーション内の分岐点の場所をパーセントで指定します。|
|*Transparency*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**単精度浮動小数点型 (Single)**|グラデーションの分岐点の色の不透明度を指定します。|
|*Index*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|グラデーションの分岐点のインデックス番号です。|
|*Brightness*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**単精度浮動小数点型 (Single)**|グラデーションの分近点の色の明るさを指定します。|



### 戻り値
なし





## 注釈
グラデーションとは、色の状態を滑らかに移行することです。このセクションのエンドポイントを分岐点と呼びます。



このメソッドは、[Insert](98aec7ed-44f9-c9b4-7a1a-e5b9a1d26d95.md) メソッドとは異なります。このメソッドでは、グラデーションの分岐点の色の明るさと透明度を指定できます。



## Related Topics

[GradientStops オブジェクト](365949f0-29b3-76e1-1163-2ac870f68f7a.md)

[GradientStops オブジェクトのメンバー](9cab316d-3302-a119-b02b-54eea372acee.md)




