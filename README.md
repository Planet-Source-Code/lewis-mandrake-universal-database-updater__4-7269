<div align="center">

## Universal Database Updater


</div>

### Description

Why I wrote thise code:

While working on a project, I realized that I was writing too many update statements. Not that it's hard, but hand coding update statements can feel like pulling teeth if you are working with large applications where you do a lot of updating. So why not create a dynamic update statement that saves time, and effort, and only needs to be written once? Okay, I admit it, I can be lazy sometimes. But this one is actually useful.

This code takes values from a querystring and uses them to update two fields within a database record. It can be used for any table in any access database.
 
### More Info
 
'The only hard coded value is the database path.

'This can be changed.

'Query String Values:

'table: Your table name

'field1: Your first field name

'field1_value: the value you want to put into your field

'field2: Your second field name

'field2_value: The value you want to put into your second field.

'where_value: the name of your primary key field

'primarykey: the value in your primary key field

'diagnostic: if diagnostic=test then it will tell

'you what has been updated

'reset: If you are not using it in diagnostic

'mode, your reset should be the URL you want to

'redirect to.

I'm guessing most intermediate programmers will be able to read this. It could be a big time saver if used properly.

'Jumping for joy.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis Mandrake](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-mandrake.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-mandrake-universal-database-updater__4-7269/archive/master.zip)

### API Declarations

mention flying monkeys.


### Source Code

```
<%
DataConnection = "Driver={Microsoft Access Driver (*.mdb)};Dbq=c:\database\data.mdb;"
If request.querystring("state")="update" then
set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = DataConnection
Command1.CommandText = "UPDATE "&request.querystring("table")&" SET "&request.querystring("field1")&"='"&request.querystring("field1_value")&"',"&request.querystring("field2")&"='"&request.querystring("field2_value")&"' WHERE "&request.querystring("where_value")&"="&request.querystring("primarykey")&""
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()
if request.querystring("diagnostic") = "test" then
Response.Write "the table <b>"&request.querystring("table")&" </b> has had the following values updated<BR>"
Response.Write "<B>"&request.querystring("field1")&" </b>has been updated to use <b>"&request.querystring("field1_value")&"</b> as it's value <BR>"
Response.Write "<B>"&request.querystring("field2")&" </b> has been updated to use <b>"&request.querystring("field2_value")&" </b>as it's value <BR>"
Response.Write "Your primary key is <B>"&request.querystring("where_value")&" </B> and it has updated where the value of that key is set to <b>"&request.querystring("primarykey")&"</b><p>"
Response.Write "This update is based on the values in the querystring. To change these values, or update a different set of fields or tables, then just change the values in the address bar above."
Response.Write "</p><p><b>Universal Querystring Updater</b><br>By <a href='http://sammoses.com'>Sam Moses</a> (c) 2002</p>"
end if
if request.querystring("diagnostic")="" then response.redirect "" &request.querystring("reset")&""
end if
%>
```

