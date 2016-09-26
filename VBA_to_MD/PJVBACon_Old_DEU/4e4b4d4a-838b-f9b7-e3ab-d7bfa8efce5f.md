
# Application.SetAutoFilter Method (Project)

Legt die Kriterien für einen AutoFilter für ein bestimmtes Feld in einer Tabellenansicht fest.


## Syntax

 _Ausdruck_. **SetAutoFilter**( ** _FieldName_**, ** _FilterType_**, ** _Test1_**, ** _Criteria1_**, ** _Operation_**, ** _Test2_**, ** _Criteria2_** )

 _Ausdruck_ Ein Ausdruck, der ein **Application** -Objekt zurückgibt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _FieldName_|Erforderlich|**String**|Der Name des Felds.|
| _FilterType_|Optional|**PjAutoFilterType**|Der Typ des Filters. Dies kann eine der  **[PjAutoFilterType](f7bd2ed9-90a1-63e9-493c-28c9c944795b.md)** -Konstanten sein. Der Standardwert ist **PjAutoFilterClear**, die den AutoFilter deaktiviert.|
| _Test1_|Optional|**String**|Gibt die Art des Vergleichs für den ersten Test an. Erfordert, dass  _FilterType_ **PjAutoFilterCustom ist** und diese _Criteria1_ gibt einen Wert an. Dies kann eine der folgenden Vergleichszeichenfolgen sein:

|**Vergleichszeichenfolge**|**Beschreibung**|
|:-----|:-----|
|"Gleich"|Der Wert von  _FieldName_ ist gleich _Criteria1_.|
|"Ungleich"|Der Wert von  _FieldName_ ist ungleich _Criteria1_.|
|"Größer als"|Der Wert von  _FieldName_ ist größer als _Criteria1_.|
|"Größer oder gleich"|Der Wert von  _FieldName_ ist größer oder gleich _Criteria1_.|
|"Kleiner als"|Der Wert von  _FieldName_ ist kleiner als _Criteria1_.|
|"Kleiner oder gleich"|Der Wert von  _FieldName_ ist kleiner oder gleich _Criteria1_.|
|"Innerhalb"|Der Wert von  _FieldName_ ist innerhalb von _Criteria1_.|
|"Nicht innerhalb"|Der Wert von  _FieldName_ ist nicht innerhalb von _Criteria1_.|
|
| _Criteria1_|Optional|**String**|Der Wert für den ersten Vergleich mit dem Wert des durch FieldName angegebenen Felds.|
| _Operation_|Optional|**String**|Die logische Operation, wenn es einen zweiten Test gibt. Der Operation-Wert kann auf And oder Or eingestellt werden.|
| _Test2_|Optional|**String**|Gibt die Art des Vergleichs für den zweiten Test an. Erfordert, dass  _FilterType_ **PjAutoFilterCustom ist** und der _Operation_ -Wert festgelegt werden muss, _Criteria2_ gibt einen Wert an. Die Zeichenfolge kann eine der in der Tabelle für Test1 Vergleiche sein.|
| _Criteria2_|Optional|**String**|Der Wert für den zweiten Vergleich mit dem Wert des durch  _FieldName_ angegebenen Felds.|

### Rückgabewert

 **Boolean**


## Hinweise

Wie die AutoFilter-Funktion aktiviert bzw. deaktiviert wird, ist unter der  **[AutoFilter](391d5a61-cba3-9e28-c448-d0befcc456c7.md)** -Methode beschrieben.


 **Hinweis**  Ein Spaltenname in einer Tabellenansicht kann einen Titel haben, der sich vom Namen des dargestellten Felds unterscheidet.


## Beispiel

Im folgenden Beispiel wird ein benutzerdefinierter AutoFilter für das Feld  **% Arbeit abgeschlossen** festgelegt.


```
Sub TestAutoFilter() 
    If Not ActiveProject.AutoFilter Then 
        Application.AutoFilter 
    End If 
 
    Application.SetAutoFilter FieldName:="% Work Complete", FilterType:=pjAutoFilterCustom, _ 
    Test1:="equals", Criteria1:="0%" 
End Sub
```

Wenn vorhanden, dass ein AutoFilter für das Feld % Arbeit abgeschlossen"festlegen ist, löscht die folgende Codezeile den AutoFilter, da der Standardwert für das optionale  _FilterType_ -Argument **PjAutoFilterClear** lautet.




```
Application.SetAutoFilter FieldName:="% Work Complete"
```

