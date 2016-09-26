
# TimeScaleValue.Clear Method (Project)

Der Wert eines Zeitskalen-Datenelements wird gelöscht.


## Syntax

 _Ausdruck_. **Clear**

 _Ausdruck_ Eine Variable, die ein **TimeScaleValue** -Objekt darstellt.


## Beispiel

Im folgenden Beispiel werden Freitage als halbe Arbeitstage festgelegt, indem eine von 08:00 Uhr bis 12:00 Uhr dauernde Schicht erstellt wird und die zweite und dritte Schicht gelöscht werden.


```
Sub HalfDayFridays() 
 With ActiveProject.Calendar.Weekdays(pjFriday) 
 .Shift1.Start = #8:00:00 AM# 
 .Shift1.Finish = #12:00:00 PM# 
 .Shift2.Clear 
 .Shift3.Clear 
 End With 
End Sub
```

