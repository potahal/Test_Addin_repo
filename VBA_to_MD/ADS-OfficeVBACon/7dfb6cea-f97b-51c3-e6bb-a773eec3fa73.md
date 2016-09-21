

---
ms.Toctitle:EncryptionProvider.Save メソッド (Office)
title:EncryptionProvider.Save メソッド (Office)
ms.ContentId:7dfb6cea-f97b-51c3-e6bb-a773eec3fa73
---
# EncryptionProvider.Save メソッド (Office)




暗号化された文書を保存します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Save**(**SessionHandle**, **EncryptionData**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **EncryptionProvider** オブジェクトを表すオブジェクト式を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*SessionHandle*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**長整数型 (Long)**|現在のセッションの ID です。|
|*EncryptionData*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**IUnknown**|暗号化情報が含まれています。|



### 戻り値
長整数型 (Long)





## 注釈
ユーザー設定のファイル暗号化をサポートする唯一の形式である Office Open XML ファイル形式にファイルを保存すると、COM アドインによってプロバイダーが呼び出されてその文書を暗号化します。ユーザー設定のファイル暗号化をサポートしない形式に保存しようとして、その保存のための適切な権限がある場合、Microsoft Office は暗号化せずに保存します。これによって、文書は暗号化やアクセス権管理をサポートしない形式にもエクスポートされます。



## Related Topics

[EncryptionProvider オブジェクト](9f5cc550-6bcb-2748-14a7-696cf8ef021b.md)

[EncryptionProvider オブジェクトのメンバー](48bed5b8-b284-4b52-4143-153ae1c751a4.md)




