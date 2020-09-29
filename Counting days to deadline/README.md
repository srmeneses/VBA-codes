### What this code should do
Calculate the number of days between today's date and the deadline when opening the file

### Code in the file

```VBA
Private Sub Workbook_Open()
    
    'First line with data because the first line contains the header
    current_line = 2
    
    'Find the last line filled with valid data
    last_line = Range("A1").End(xlDown).Row
    
    'Iterate through all lines that contain valid data
    While current_line <= last_line
        
        'Copy the deadline from the current row
        deadline = Range("C" & current_line)
        
        'Calculate the number of days between today's date and the deadline
        remaining_days = DateDiff("d", Now, deadline)
        
        'Copy the remaining_days to the current row
        Range("D" & current_line) = remaining_days
        
        'Go to the next line
        current_line = current_line + 1
    Wend
    
End Sub
```

### Useful links
* [Workbook_Open](https://docs.microsoft.com/pt-br/office/vba/api/excel.workbook.open)
* [DateDiff](https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/datediff-function)
* [Range](https://docs.microsoft.com/pt-br/office/vba/api/excel.range(object))
* [While - Wend](https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/whilewend-statement)
* [End(xlDown)](https://docs.microsoft.com/pt-br/office/vba/api/excel.range.end)

