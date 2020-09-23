<div align="center">

## Dynamically Create Databases \(\.MDB's\) in code


</div>

### Description

This code creates a Microsoft Access MDB dynamically.
 
### More Info
 
This sample code will create a database in the temp directory with the following fields:

fldForeName

fldSurname

fldDOB

fldFurtherDetails


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kenny Sendel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kenny-sendel.md)
**Level**          |Unknown
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kenny-sendel-dynamically-create-databases-mdb-s-in-code__1-1253/archive/master.zip)

### API Declarations

NONE


### Source Code

```
''
'' PUT THIS BEHIND A COMMAND BUTTON TO TEST
''
' Declarations
Dim tdExample      As TableDef
Dim fldForeName     As Field
Dim fldSurname     As Field
Dim fldDOB       As Field
Dim fldFurtherDetails  As Field
Dim dbDatabase     As Database
Dim sNewDBPathAndName  As String
' Set the new database path and name in string (using time:seconds for some randomality
sNewDBPathAndName = "c:\temp\NewDB" & Right$(Time, 2) & ".mdb"
' Create a new .MDB file (empty at creation point!)
Set dbDatabase = CreateDatabase(sNewDBPathAndName, dbLangGeneral, dbEncrypt)
' Create new TableDef (table called 'Example')
Set tdExample = dbDatabase.CreateTableDef("Example")
' Add fields to tdfTitleDetail.
Set fldForeName = tdExample.CreateField("Fore_Name", dbText, 20)
Set fldSurname = tdExample.CreateField("Surname", dbText, 20)
Set fldDOB = tdExample.CreateField("DOB", dbDate)
Set fldFurtherDetails = tdExample.CreateField("Further_Details", dbMemo)
' Append the field objects to the TableDef
tdExample.Fields.Append fldForeName
tdExample.Fields.Append fldSurname
tdExample.Fields.Append fldDOB
tdExample.Fields.Append fldFurtherDetails
' Save TableDef definition by appending it to TableDefs collection.
dbDatabase.TableDefs.Append tdExample
MsgBox "New .MDB Created - '" & sNewDBPathAndName & "'", vbInformation
' Now look for the new .MDB using File Manager!
```

