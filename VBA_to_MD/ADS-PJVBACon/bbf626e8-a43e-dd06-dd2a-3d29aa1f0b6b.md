

---
ms.Toctitle:Project.DeliverablesGetByProject メソッド (Project)
title:Project.DeliverablesGetByProject メソッド (Project)
ms.ContentId:bbf626e8-a43e-dd06-dd2a-3d29aa1f0b6b
---
# Project.DeliverablesGetByProject メソッド (Project)




取得したオブジェクトの XML メンバーで指定されたエンタープライズ プロジェクト用のすべての成果物の一覧を取得します。Project Professional のみ。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**DeliverablesGetByProject**(**ProjectGuid**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Project** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*ProjectGuid*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|エンタープライズ プロジェクトの GUID です。|



### 戻り値
**オブジェクト**





## Remarks
VBA を使用して **DeliverablesGetByProject** 結果オブジェクトの **XML** メンバーを処理するには、複雑で直感的に理解できないコードが必要です。Project Server と SharePoint の機能を使用する場合は、UNRESOLVED_TOKEN_VAL(vsdev11short) で Office と SharePoint の開発ツールを使用して、Project のアドインを作成することをお勧めします。XML を処理する最も簡単なアプローチは、UNRESOLVED_TOKEN_VAL(dotnetfw40short) で LINQ to XML メソッドを使用する方法です。



## 次の使用例では、テーブルからレコードを削除できないようにします。
以下の例では、Simple という名前のエンタープライズ プロジェクトに、M1 という名前のマイルストーンが設定された成果物が含まれています。成果物一覧が含まれたプロジェクト サイトの URL は `http://ServerName/PWA/Simple` です。**TestDeliverables** マクロによって、XML 結果の部分が含まれるメッセージ ボックスを表示します。



**GetServerProjectGuid** メソッドで取得した **projectGuid** 値には、たとえば "{1b14e65c-5601-4565-acb9-3822078a17fb}" のような GUID を囲むかっこが含まれています。GUID 値はかっこ付きまたはかっこなしのいずれかで使用できます。

```vba
Option Explicit 
 
Sub TestDeliverables() 
    Dim projectGuid As String 
    Dim ds As Object 
 
    projectGuid = ActiveProject.GetServerProjectGuid 
 
    ' Optional: Removing the braces on the GUID value makes no difference. 
    ' projectGuid = Mid(projectGuid, 2, 36) 
 
    Set ds = ActiveProject.DeliverablesGetByProject(projectGuid) 
 
    MsgBox ds.XML 
 
    Debug.Print ds.XML 
End Sub
```


>[!NOTE]
>**ds** などの **Object** 型変数のメンバーを検出するには、オブジェクトにウォッチを設定して、オブジェクトに値を割り当てた後はブレークポイントを設定します。[**ウォッチ**] ウィンドウで変数を展開すると、**XML** メンバーを表示できます。





メッセージ ボックスには、この例の XML の結果である合計17,295 文字のうち、最初の1,024 文字のみ表示されます。以下の XML の結果では、属性はそれぞれ別の行に分かれていますが、実際の XML の結果はすべて 1 行に記述されているので、VBE の [**イミディエイト**] ウィンドウに結果を印刷できるかどうかを確認できます。この例では、コンテンツのほとんどを構成する XML スキーマは表示されていません。



**ows_** フィールドは SharePoint 一覧で定義されています。抽出する可能性のあるフィールドには、**deliverableUid**、**workspaceUri**、**linkedTaskUid** (Project Server のタスクの GUID)、**ows_LinkTitle** (成果物を含むタスクの名前)、**ows_Created**、**ows_Modified**、**ows_Author**、**ows_CommitmentStart**、および **ows_CommitmentFinish** が含まれます。

```xml
<DeliverableMasterDocument> 
 <Deliverables> 
 <Deliverable deliverableUid="6f8cb9a5-d9b8-496d-af90-1e88dc57f46a" projectUid="1b14e65c-5601-4565-acb9-3822078a17fb" 
 type="1" tpId="1" workspaceUri="http://ServerName/PWA/Simple" workspaceName="PWA/Simple" workspaceVServerUri="http://ServerName" 
 listUid="168a6e6f-6993-4315-a593-7ffa21683e57" state="1"> 
 <Client linkedTaskUid="d3eaf532-9ab9-4eb2-8f85-fd41a1b5db0c" ows_ID="1" 
 ows_ContentTypeId="0x010074416DB49FB844B99C763FA7171E7D1F00001031A192BFCA4D83CA160D2BCAB735" 
 ows_ContentType="Project Site Deliverable" ows_Title="M1" ows_Modified="2010-02-19 13:30:19" 
 ows_Created="2010-02-19 13:29:45" ows_Author="1073741823;#System Account" 
 ows_Editor="1073741823;#System Account" ows_owshiddenversion="2" ows_WorkflowVersion="1" 
 ows__UIVersion="512" ows__UIVersionString="1.0" ows_Attachments="0" ows__ModerationStatus="0" 
 ows_LinkTitleNoMenu="M1" ows_LinkTitle="M1" ows_LinkTitle2="M1" ows_SelectTitle="1" 
 ows_Order="100.000000000000" ows_GUID="{FFA3E0F9-DBB4-44B6-B09D-1C2AB7A9EF92}" 
 ows_FileRef="1;#PWA/Simple/Lists/Deliverables/1_.000" ows_FileDirRef="1;#PWA/Simple/Lists/Deliverables" 
 ows_Last_x0020_Modified="1;#2010-02-19 13:29:45" ows_Created_x0020_Date="1;#2010-02-19 13:29:45" 
 ows_FSObjType="1;#0" ows_SortBehavior="1;#0" ows_PermMask="0x7fffffffffffffff" ows_FileLeafRef="1;#1_.000" 
 ows_UniqueId="1;#{29AF34EA-EA27-48C7-80A6-83B0A95DB9BD}" ows_ProgId="1;#" 
 ows_ScopeId="1;#{73C1A12E-DBA2-4BE2-87EE-1FF5EF1494DD}" ows__EditMenuTableStart="1_.000" 
 ows__EditMenuTableStart2="1" ows__EditMenuTableEnd="1" ows_LinkFilenameNoMenu="1_.000" 
 ows_LinkFilename="1_.000" ows_LinkFilename2="1_.000" ows_ServerUrl="/PWA/Simple/Lists/Deliverables/1_.000" 
 ows_EncodedAbsUrl="http://jc2vm1/PWA/Simple/Lists/Deliverables/1_.000" ows_BaseName="1_" ows_MetaInfo="1;#" 
 ows__Level="1" ows__IsCurrentVersion="1" ows_ItemChildCount="1;#0" ows_FolderChildCount="1;#0" 
 ows_CommitmentStart="2010-02-02 00:00:00" ows_CommitmentFinish="2010-02-02 00:00:00" ows_SuppressCreateEvent="1"/> 
 </Deliverable> 
 </Deliverables> 
 <Schemas> 
 <Schema . . . 
 . . . > 
 <Fields> 
 <Field . . . /> 
 . . . 
 </Fields> 
 </Schema> 
 </Schemas> 
</DeliverableMasterDocument>
```





