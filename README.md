<div align="center">

## Return a Group of Random Records


</div>

### Description

To return a group of Random Records from a database. For example, a group of random questions for a quiz/test.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Larry Boggs](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/larry-boggs.md)
**Level**          |Beginner
**User Rating**    |4.2 (25 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/larry-boggs-return-a-group-of-random-records__4-6201/archive/master.zip)





### Source Code

<!--random.asp-->
<!--Copyright (c) 1999 by Larry L. Boggs. All rights reserved.-->
<!--Generate a random recordset from an Access database-->
<html>
<head>
<title>Random Recordset</title>
<meta name="author" content="Larry Boggs">
<meta name="email" content="lboggs@i1.net">
<meta name="date" content="5/21/00">
<meta name="description" content="Return a Random Group of Records">
<meta name="keywords" content="access, random, group, records, recordset">
</head>
<body>
<center><b><h1><p>
Return a Group of Random Records
</p></h1></b></center>
<p>
While working on a web based competency-testing application I needed a way to return not just ONE random record but a group of random records. I searched the net high and low for a couple of months
trying to find something that would allow me do this. I eventually hunkered down and came up with my own way of doing this.
</p>
<p>
First comes the SQL statement to return the set of records you will pick your Random records from:
</p>
<p>
<table width=95% border=0>
<tr><td bgcolor=#cccccc width=100%>
<code><pre>
&lt;%
  strConnection="driver={Microsoft Access Driver (*.mdb)};dbq=" & server.mappath("/testdb.mdb")
  strSQL = "SELECT id FROM tblQuestions"
  set objConn = Server.CreateObject("ADODB.Connection")
  Set objRst = Server.CreateObject("ADODB.Recordset")
  objConn.Open strConnection
  set objRst.ActiveConnection = objConn
  objRst.LockType = adLockOptimistic
  objRst.CursorType = adOpenKeySet
  objRst.Open strSQL
%&gt;</pre></code>
</td></tr>
</table>
</p>
<p>
Next, set the upper limit of the Randomize function by setting the variable rndMax equal to the RecordCount.
</p>
<p>
<table width=95% border=0>
<tr><td bgcolor=#cccccc width=100%>
<code><pre>
&lt;%
  objRst.MoveLast
  cnt = objRst.RecordCount
  cnt1 = cnt
  rndMax = cnt
%&gt;</pre></code>
</td></tr>
</table>
</p>
<p>
Next, set the number of records returned to either the number of questions they asked for or equal to the RecordCount.
</p>
<p>
<table width=95% border=0>
<tr><td bgcolor=#cccccc width=100%>
<code><pre>
&lt;%
  If CInt(Request.Form("maxNumber")) < cnt Then
	cnt1 = CInt(Request.Form("maxNumber"))
  End If
%&gt;</pre></code>
</td></tr>
</table>
</p>
<p>
Now we want to return a Random number. Check if the variable &#8220;str1&#8221; already contains that number. If so then that number is skipped
and it loops again returning another Random record number. This ensures that NO values are repeated. If not then plug that number into
the &#8220;str1&#8221; variable so we will know that that number has already been used the next time through the loop. If the random number is not
contained within the &#8220;str1&#8221; variable then the value of the &#8220;ID&#8221; field is returned and plugged into the &#8220;str&#8221; variable. This loops until the
appropriate number of values have been plugged into the &#8220;str&#8221; variable.
</p>
<p>
<table width=95% border=0>
<tr><td bgcolor=#cccccc width=100%>
<code><pre>
&lt;%
  str = ","
  str1 = ","
  Do Until cnt1 = 0
    Randomize
    RndNumber = Int(Rnd * rndMax)
    If (InStr(1, str1, "," & RndNumber & "," ) = 0) Then
	  str1 = str1 & RndNumber & ","
	  cnt1 = cnt1 - 1
	  objRst.MoveFirst
	  objRst.Move RndNumber
      str = str & objRst("id") & ","
	End If
  Loop
%&gt;</pre></code>
</td></tr>
</table>
</p>
<p>
Now we have a variable, (str), that contains a comma-delimited list of values from the &#8220;ID&#8221; field. Now, just reference the comma-
delimited string contained within the &#8220;str&#8221; variable in your SQL statement:
</p>
<p>
<table width=95% border=0>
<tr><td bgcolor=#cccccc width=100%>
<code><pre>
&lt;%
   sql = "SELECT * FROM tblQuestions WHERE (((InStr(1,'" & str & "',(',' & [id] & ',')))<>0)) "
%&gt;</pre></code>
</td></tr>
</table>
</p>
<p>
This will return your Random set of records!
</p>
<p>
Here's the whole thing:
</p>
<p>
<table width=95% border=0>
<tr><td bgcolor=#cccccc width=100%>
<code><pre>
&lt;%
<!--Generate a random recordset from an Access database-->
<!--#include virtual="/adovbs.inc"-->
&lt;%
  Dim objConn
  Dim objRst
  Dim strSQL
  Dim strConnection
  Dim str
  Dim str1
  Dim cnt
  Dim cnt1
  Dim rndMax
  Dim RndNumber
  strConnection="driver={Microsoft Access Driver (*.mdb)};dbq=" & server.mappath("/testdb.mdb")
  strSQL = "SELECT id FROM tblQuestions"
  set objConn = Server.CreateObject("ADODB.Connection")
  Set objRst = Server.CreateObject("ADODB.Recordset")
  objConn.Open strConnection
  set objRst.ActiveConnection = objConn
  objRst.LockType = adLockOptimistic
  objRst.CursorType = adOpenKeySet
  objRst.Open strSQL
  objRst.MoveLast
  cnt = objRst.RecordCount
  cnt1 = cnt
  rndMax = cnt
  If CInt(Request.Form("maxNumber")) < cnt Then
	cnt1 = CInt(Request.Form("maxNumber"))
  End If
  str = ","
  str1 = ","
  Do Until cnt1 = 0
    Randomize
    RndNumber = Int(Rnd * rndMax)
    If (InStr(1, str1, "," & RndNumber & "," ) = 0) Then
	  str1 = str1 & RndNumber & ","
	  cnt1 = cnt1 - 1
	  objRst.MoveFirst
	  objRst.Move RndNumber
      str = str & objRst("id") & ","
	End If
  Loop
  objRst.Close
  Set objRst = Nothing
  sql = "SELECT * FROM tblQuestions WHERE (((InStr(1,'" & str & "',(',' & [id] & ',')))<>0)) "
  Set objRst = Server.CreateObject("ADODB.Recordset")
  set objRst.ActiveConnection = objConn
  objRst.LockType = adLockOptimistic
  objRst.CursorType = adOpenKeySet
  objRst.Open sql
%&gt;
...DISPLAY THE RECORDS RETURNED...
&lt;%
objRst.Close
Set objRst = Nothing
objConn.Close
Set objConn = Nothing
%&gt;
</pre></code>
</td></tr>
</table>
</p>
<p>
I'd be interested in hearing from anyone that builds upon this and/or how they put it to use!<br>
<br>
See Ya!<br>
<b><a HREF="mailto:lboggs@i1.net">
Larry Boggs</a></b>
</p>
</body>
</html>

