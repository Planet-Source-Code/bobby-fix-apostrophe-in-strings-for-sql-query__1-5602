<div align="center">

## Fix Apostrophe in strings for SQL query


</div>

### Description

Fix Apostrophe in strings for SQL query
 
### More Info
 
String as input to the function

Returns the modified string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bobby](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bobby.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bobby-fix-apostrophe-in-strings-for-sql-query__1-5602/archive/master.zip)





### Source Code

```
Function Fix_Apostrophe(ByVal S As String) As String
  Dim i As Integer, ch As String, Ret As String
  If IsNull(S) Then Exit Function
  Ret = ""
  For i = 1 To Len(S)
    ch = Mid$(S, i, 1)  ' the current charcater
    Ret = Ret & ch
    ' If the character is a single quote add a second one.
    If ch = "'" Then
     Ret = Ret & ch
    End If
  Next
  Fix_Apostrophe = Ret
End Function
```

