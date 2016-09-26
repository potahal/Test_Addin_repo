
# PostItem.ReadComplete-Ereignis (Outlook)
Tritt auf, wenn Outlook abgeschlossen ist, lesen die Eigenschaften des Elements.

## Versionsinformationen

Hinzugefügte Version: Outlook 2013


## Syntax

 _Ausdruck_. **ReadComplete** _(Cancel)_

 _Ausdruck_ Eine Variable, die ein PostItem **PostItem**-Objekt darstellt


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|||||
| _Cancel_|Erforderlich|**Boolean**|(In VBScript nicht verwendet).  **False** Wenn das Ereignis auftritt. Wenn die Ereignisprozedur dieses Argument auf **True**festgelegt wird, der Lesevorgang nicht abgeschlossen und das Element wird im Lesebereich oder im Inspektor nicht angezeigt.|

## Hinweise

Das  **ReadComplete** -Ereignis tritt nach dem[BeforeRead](26a64e4e-a48e-84e8-4fea-70913a8f170f.md) -Ereignis und vor dem[Lesen](404c9b17-c5b6-a802-420a-f8fd279b5f9b.md) -Ereignis für das Element.

Um zu bestimmen, wann das Element aus dem Speicher entfernt wird, verwenden Sie das [Unload](42dea931-c3dd-b8ff-5ace-0744b17650e0.md)-Ereignis.

Das  **ReadComplete** -Ereignis entspricht dem Exchange Client Extensions (ECE)-Ereignis **IExchExtMessageEvents::OnReadComplete**.


## Siehe auch


#### Konzepte


[PostItem-Objekt](de44065d-4e93-315a-279f-7b92f09c0465.md)
#### Weitere Ressourcen


[PostItem-Objektmember](http://msdn.microsoft.com/library/5b150db1-c96d-0721-ec36-d5b5ebc20fd8%28Office.15%29.aspx)