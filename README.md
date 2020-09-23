<div align="center">

## Beginning & End of Month for Processing \- Dates


</div>

### Description

Determine Beginning & End of Month for Processing
 
### More Info
 
Date derived from system

Beginning and end dates of the previous month for financial calculations


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sean Wilkins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sean-wilkins.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sean-wilkins-beginning-end-of-month-for-processing-dates__1-34774/archive/master.zip)

### API Declarations

```
Public sBegDtTmp As String * 10    'beginning date for calcs
Public sEndDtTmp As String * 10    'ending date for calcs
```


### Source Code

```
sEndDtTmp = Format$((DateValue(Now) - Day(Now)), "mm/dd/yyyy")
sBegDtTmp = Format$(DateSerial(Year(sEndDtTmp), Month(sEndDtTmp), 1), "mm/dd/yyyy")
```

