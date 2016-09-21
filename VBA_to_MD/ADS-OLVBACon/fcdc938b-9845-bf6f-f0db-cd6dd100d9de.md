

---
ms.Toctitle:FormRegion.DisplayName プロパティ (Outlook)(機械翻訳)
title:FormRegion.DisplayName プロパティ (Outlook)(機械翻訳)
ms.ContentId:fcdc938b-9845-bf6f-f0db-cd6dd100d9de
---
# FormRegion.DisplayName プロパティ (Outlook)(機械翻訳)




フォーム領域の表示名を表す**文字列**を返します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**DisplayName**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **FormRegion** オブジェクトを表す変数です。



## 注釈
表示名は、フォーム領域では省略可能です。対応するフォーム領域マニフェスト XML ファイルで < formRegionName > タグの値を定義した場合、この値は**DisplayName**プロパティの値にマップされます。フォーム領域の XML スキーマの詳細については、 [MSDN ライブラリ](http://msdn.microsoft.com/library)で Microsoft Outlook 2010 の XML スキーマ リファレンスを参照してください。



**DisplayName**プロパティの値が実行時に、個別フォーム領域のリボンの [**表示**] タブで、または隣接するフォーム領域のヘッダーに表示されます。 既定のロケールを使用し、対応するフォーム領域マニフェスト XML ファイルで < stringOverride > タグで無効にできます。文字列は、大文字とその最大の長さは、256 文字です。



## Related Topics

[FormRegion オブジェクト](3a0b83eb-4076-9cb3-86a9-68f9e44df89f.md)

[FormRegion オブジェクトのメンバー](eb4ff750-2911-8f8d-2ef0-c3f5e7adf4e0.md)




