

---
ms.Toctitle:FormRegionStartup.GetFormRegionIcon メソッド (Outlook)(機械翻訳)
title:FormRegionStartup.GetFormRegionIcon メソッド (Outlook)(機械翻訳)
ms.ContentId:c1c0bd3f-3fae-8e9b-d579-58d609bbaa4e
---
# FormRegionStartup.GetFormRegionIcon メソッド (Outlook)(機械翻訳)




フォーム領域の特定の種類のアイコンに対して表示されるアイコン イメージを取得します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GetFormRegionIcon**(**FormRegionName**, **LCID**, **Icon**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **FormRegionStartup** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*FormRegionName*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|フォーム領域を Windows レジストリに登録するときに使用されたフォーム領域の名前。|
|*LCID*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**長整数型**|Outlook が現在使用している言語を表すロケール ID。この値は、この言語に対応するフォーム領域のローカライズ文字列を取得するために使用します。|
|*Icon*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**OlFormRegionIcon**|アイコンの種類を識別する定数を指定します。|



### 戻り値
オリジナルのイメージ ファイルのバイト数を表すバイト配列または**IPictureDisp**オブジェクトのいずれかであるバリアント型です。





## 注釈
このメソッドはアドインによって実装され、Outlook によって呼び出されます。**FormRegionStartup** インターフェイスの一部として、このメソッドと **GetFormRegionManifest** メソッドには、アドインがフォーム領域を登録し、フォーム領域の XML マニフェストとアイコンを Outlook に提供するメカニズムが用意されています。



フォーム領域のアイコンを提供するアドインの場合は、Windows レジストリでフォーム領域を登録するときにアドインの ProgID を指定します。フォーム領域を登録する方法の詳細については、 [Windows レジストリでフォーム領域を指定する](0de3fcb1-b357-8300-c943-9a5a788d4976.md)を参照してください。アドインでは、 **GetFormRegionManifest** 、 **FormRegionStartup**インターフェイスの**GetFormRegionIcon**メソッドを実装しなければなりません。



**アイコン**要素の下で、フォーム領域の XML マニフェストには、カスタム アイコンを使用するには子要素のそれぞれの値`addin`を指定します。 **GetFormRegionIcon**がカスタム アイコンのイメージを取得するとき Outlook アイコンの種類*のアイコン*を引数として、そのような**GetFormRegionIcon**を実装します。既定のアイコンを表示するように Outlook を実行する場合に、 **null** (**何も**Visual Basic で) のアイコンの種類が返されますように、 **GetFormRegionIcon**を実装します。**GetFormRegionIcon**を**null** (**Nothing**で Visual Basic) も返す必要があります*アイコン*は、 **olFormRegionIconDefault**とします。



Outlook の起動時、Windows レジストリからフォーム領域の一覧を読み取りし、フォーム領域に関連付けられているデータをキャッシュします。ProgID を持つフォーム領域を登録している場合 Outlook は**アイコン**の要素の子要素の値として`addin`を持つ XML マニフェスト内の任意のアイコンの**GetFormRegionIcon**の実装を呼び出すことで、対応するアドインを並び替え。呼び出さないことを場合は、Windows レジストリ内の ProgID を指定しないと、Outlook は、 **GetFormRegionManifest**メソッドと**GetFormRegionIcon**メソッドに注意してください。



## Related Topics

[FormRegionStartup オブジェクトのメンバー](c45b60b8-5d7e-d84b-a60e-ffcb54c25569.md)

[FormRegionStartup インターフェイス](948ea6b7-2962-57e7-618d-fa0977b65651.md)




