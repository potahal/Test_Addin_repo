
# Months Object (Project)

Enthält eine Auflistung von  **[Month](5ee32f12-72aa-fa16-ead2-97949005cd7c.md)** -Objekten.


## Bemerkungen

Verwenden Sie  **Months** ( _Index_ ), wobei _Index_ der Monatsindex, der Monatsname oder die **PjMonth** -Konstante ist, um ein einzelnes **Month** -Objekt zurückzugeben.


## Beispiel

 **Verwenden des Months-Auflistungsobjekts**

Im folgenden Beispiel wird ermittelt die Anzahl von Arbeitstagen in jeden Monat des 2012 für jede ausgewählte Ressource.




```
Dim R As Resource 
Dim D As Integer, M As Integer, WorkingDays As Integer 
 
For Each R In ActiveSelection.Resources() 
    WorkingDays = 0 

    With R.Calendar.Years(2012) 
        For M = 1 To .Months.Count 
            WorkingDays = 0 
            For D = 1 To .Months(M).Days.Count 
                If .Months(M).Days(D).Working = True Then 
                    WorkingDays = WorkingDays + 1 
                End If 
            Next D 

            MsgBox "There are " &amp; WorkingDays &amp; " working days in " &amp; _
                .Months(M).Name &amp; " for " &amp; R.Name &amp; "." 
        Next M 
    End With 
Next R
```

 **Verwenden der Months-Auflistung**

Verwenden Sie die  **[Months](615a4f5c-bda7-f684-1c29-d8003badf3a8.md)** -Eigenschaft, um eine **Months** -Auflistung zurückzugeben. Im folgenden Beispiel wird ermittelt die Anzahl der Monate in 2012.




```
ActiveProject.Calendar.Years(2012).Months.Count
```


## Siehe auch


#### Konzepte


[Projektobjektmodell](900b167b-88ec-ea88-15b7-27bb90c22ac6.md)