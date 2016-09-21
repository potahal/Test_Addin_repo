

---
ms.Toctitle:Application.OpenXML メソッド (Project)
title:Application.OpenXML メソッド (Project)
ms.ContentId:dcf3dd0e-78ec-b95c-b890-dca5507acd92
---
# Application.OpenXML メソッド (Project)




XML 文字列からプロジェクトを開きます。

## 構文
UNRESOLVED_TOKEN_VAL(offexpression).**OpenXML**(**XML**)



UNRESOLVED_TOKEN_VAL(offexpression)**Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*XML*|UNRESOLVED_TOKEN_VAL(offrequired)|**文字列型 (String)**|Project XML スキーマに準拠する有効な Project XML 文字列を含む文字列を指定します。|



### 戻り値
**長整数型 (Long)**





## 注釈
Project XML スキーマは、UNRESOLVED_TOKEN_VAL(pjgenericshort) SDK に mspdi_pj15.xsd ファイルとして含まれています。XML ファイルを作成するには、プロジェクトを XML に保存して編集します。プログラムで XML 文字列を作成する場合は、**OpenXML** メソッドで使用する前にその文字列がスキーマに準拠していることを確認する必要があります。



**OpenXML** メソッドは、成功すると 0 を返します。

>[!NOTE]
>有効な Project XML ファイルを開くには、 **FileOpenEx** メソッドも使用できます。**OpenXML** メソッドは、主に XML 文字列を使用してプロジェクトを開くよう設計されています。





## 例
次の使用例は、プロジェクトを XML として保存し、そのファイルを編集することで作成した OneTaskEdited.xml というファイルを開き、既定値を削除します。この例では、Microsoft Scripting Runtime ライブラリ (scrrun.dll) を参照する必要があります。

```vba
Sub ImportXMLProject() 
    ' Requires reference to the Microsoft Scripting Runtime library (scrrun.dll). 
    Dim txtStream As TextStream 
    Dim fileName As String 
    Dim xmlContents As String 
    Dim fsObject As FileSystemObject 
 
    fileName = "C:\Project\VBA\Samples\OneTaskEdited.xml" 
    Set fsObject = CreateObject("Scripting.FileSystemObject") 
 
    If Not fsObject.FileExists(fileName) Then 
        MsgBox "The file does not exist: " & vbCrLf & fileName 
    Else 
        ' Open a text stream. 
        Set txtStream = fsObject.OpenTextFile(fileName:=fileName, IOMode:=ForReading) 
 
        xmlContents = txtStream.ReadAll 
        Application.OpenXML(xmlContents) 
        txtStream.Close 
    End If 
End Sub
```





