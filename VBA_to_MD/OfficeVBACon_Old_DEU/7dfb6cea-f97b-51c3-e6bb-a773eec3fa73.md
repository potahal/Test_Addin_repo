
# EncryptionProvider.Save-Methode (Office)

Speichert ein verschlüsseltes Dokument.


## Syntax

 _Ausdruck_. **Save**( ** _SessionHandle_**, ** _EncryptionData_** )

 _Ausdruck_ Ein Ausdruck, der ein **EncryptionProvider** -Objekt zurückgibt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Erforderlich|**Long**|Die ID der aktuellen Sitzung.|
| _EncryptionData_|Erforderlich|**IUnknown**|Enthält die Verschlüsselungsinformationen.|

### Rückgabewert

Long


## Bemerkungen

Wenn Sie eine Datei im Office Open XML-Dateiformat speichern (das einzige Format, das die benutzerdefinierte Dateiverschlüsselung unterstützt), wird der Anbieter vom COM-Add-In zum Verschlüsseln des Dokuments aufgerufen. Wenn Sie versuchen, das Dokument in einem Format zu speichern, das die benutzerdefinierte Dateiverschlüsselung nicht unterstützt und Sie über die entsprechenden Rechte verfügen, wird das Dokument von Microsoft Office unverschlüsselt gespeichert. Auf diese Weise können Dokumente in Formate exportiert werden, die die Verschlüsselung oder die Rechteverwaltung nicht unterstützen.


## Siehe auch


#### Konzepte


[EncryptionProvider-Objekts](9f5cc550-6bcb-2748-14a7-696cf8ef021b.md)
#### Weitere Ressourcen


[Elemente des EncryptionProvider-Objekts](http://msdn.microsoft.com/library/48bed5b8-b284-4b52-4143-153ae1c751a4%28Office.15%29.aspx)