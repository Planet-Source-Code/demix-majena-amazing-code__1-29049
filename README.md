<div align="center">

## Amazing Code


</div>

### Description

Speed deprivation at its best
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Demix Majena](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/demix-majena.md)
**Level**          |Advanced
**User Rating**    |2.5 (10 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/demix-majena-amazing-code__1-29049/archive/master.zip)





### Source Code

```
Sub Sleep(ByVal MillaSec As Long, Optional ByVal Freeze As Boolean = False)
  Dim tStart#, Tmr#
  tStart = Timer
  While Tmr < (MillaSec / 1000)
    Tmr = Timer - tStart
    If Freeze = False Then DoEvents
  Wend
End Sub
```

