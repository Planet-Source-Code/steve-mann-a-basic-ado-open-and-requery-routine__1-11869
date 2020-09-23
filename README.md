<div align="center">

## A basic ADO Open and Requery routine


</div>

### Description

New to ADO? Worried about ADO? This little subroutine gets round the problems of opening ADO recordsets. Please look elsewhere on this site for info on opening the database itself.

If you open a Recordset in your code, ADO expects you to close it before re-opening. But if it's not open you can't close it... (Oh My!).

Here's my solution. It requires a public ADODB.Connection - I call it gCn.

The routine will open a new recordset (compatible with Janus GridEx), or refresh it if it's open or if the SQL has changed. Optional ReadOnly argument.
 
### More Info
 
rs - an ADO recordset (eg Dim rsMine as New ADODB.Recordset)

szSource - (eg "select * from customers")

Optional - bReadOnly (True for read-only)

Assumes your public ADO Connection object is called gCn, and that the supplied szSource is a valid SQL statement.

Sets the supplied Recordset.

If you pass invalid SQL, you'll get an error. Ctrl+Break and you'll be ready to F8 out and see where you went wrong.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Mann](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-mann.md)
**Level**          |Advanced
**User Rating**    |3.8 (46 globes from 12 users)
**Compatibility**  |VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-mann-a-basic-ado-open-and-requery-routine__1-11869/archive/master.zip)





### Source Code

```
Public Sub ADO_OpenRs(rs As Recordset, szSource$, Optional bReadOnly = False)
' Open or Requery a Recordset.
On Error GoTo lab_Err
If rs.State = adStateClosed Or rs.Source <> szSource Then
 If rs.State <> adStateClosed Then rs.Close
 rs.Open szSource, gCn, adOpenStatic, IIf(bReadOnly, adLockReadOnly, adLockOptimistic)
Else
 rs.Requery
End If
lab_Exit:
 Exit Sub
lab_Err:
 MsgBox Err.Description
 GoTo lab_Exit
End Sub
```

