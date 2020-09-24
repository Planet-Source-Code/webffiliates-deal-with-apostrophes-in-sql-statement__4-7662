<div align="center">

## Deal With Apostrophes in SQL Statement


</div>

### Description

This code allows you to input strings with apostrophes into a database.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[WebFFiliates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/webffiliates.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/webffiliates-deal-with-apostrophes-in-sql-statement__4-7662/archive/master.zip)





### Source Code

```
<%
'Get the data from the form
Dim strName
'Let's take the string and replace all single apostrophes with double apostrophes
strCompanyName = Replace(Request.form("NAME"), "'", "''")
 'Connect to the Database
Dim strSQL
strSQL = "INSERT YourTable (Name) VALUES ('" &_ strName & "')"
%>
```

