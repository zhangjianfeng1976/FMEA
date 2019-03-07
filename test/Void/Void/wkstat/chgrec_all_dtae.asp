<!--#Include file="includes/db.asp"-->
<% 
' 要调用的函数声明 '
'根据年份及月份得到每月的总天数 '
Function GetDaysInMonth(iMonth, iYear) 
Select Case iMonth 
Case 1, 3, 5, 7, 8, 10, 12 
GetDaysInMonth = 31 
Case 4, 6, 9, 11 
GetDaysInMonth = 30 
Case 2 
If IsDate("February 29, " & iYear) Then 
GetDaysInMonth = 29 
Else 
GetDaysInMonth = 28 
End If 
End Select 
End Function 
'得到一个月开始的日期. '
Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth) 
Dim dTemp 
dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth) 
GetWeekdayMonthStartsOn = WeekDay(dTemp) 
End Function 
'得到当前一个月的上一个月. '
Function SubtractOneMonth(dDate) 
SubtractOneMonth = DateAdd("m", -1, dDate) 
End Function 
'得到当前一个月的下一个月. '
Function AddOneMonth(dDate) 
AddOneMonth = DateAdd("m", 1, dDate) 
End Function 
'得到当前一年的上一年. '
Function SubtractOneYear(dDate) 
SubtractOneYear = DateAdd("yyyy", -1, dDate) 
End Function 
'得到当前一年的下一年.' 
Function AddOneYear(dDate) 
AddOneYear = DateAdd("yyyy", 1, dDate) 
End Function 
' 函数声明结束' 

Dim dDate ' 日历显示的日期 '
Dim iDOW ' 每一月开始的日期 '
Dim iCurrent ' 当前日期 '
Dim iPosition ' 表格中的当前位置 '


' 得到选择的日期并检查日期的合法性 '
If IsDate(Request("date")) Then 
dDate = CDate(Request("date")) 
Else 
If IsDate(Request("year") & "-" & Request("month") & "-" & Request("day")) Then 
dDate = CDate(Request("year") & "-" & Request("month") & "-" & Request("day")) 
Else 
dDate = Date() 

If Len(Request("month")) <> 0 Or Len(Request("day")) <> 0 Or Len(Request("year")) <> 0 Or Len(Request("date")) <> 0 Then 
Response.Write "您所选择的日期格式不正确，系统会使用当前日期.<BR><BR>" 
End If 

End If 
End If 

'得到日期后我们先得到这个月的天数及这个月的起始日期. '
iDIM = GetDaysInMonth(Month(dDate), Year(dDate)) 
iDOW = GetWeekdayMonthStartsOn(dDate) 

%> 
<!--------------------------------------------------------------------------------------------->
<%

if trim(Request("date"))="" then 
aa = date()
else
aa = trim(request("date"))
end if

sql="SELECT  * FROM v_wdairy where 日期 = '"&aa&"'"

Call openDB()
rs.open sql,conn,1,1
If not rs.eof then
  rs.pagesize = 500
  pagecount=rs.PageCount 
  page=int(request ("page"))
  if page<=0 then page=1
  if request("page")="" then page=1
  rs.AbsolutePage=page 
End if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>人员工位管理</title>
<link href="Includes/Style.css" rel="stylesheet" type="text/css">
</head>

<body>	
<form action="chgrec_all_dtae.asp">
<table width="100%" border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr><td height="35" align='center' bgcolor="#E6F7FF" colspan='3'>
<b>当日替位一览表[<%= aa %>]</b>
</td></tr>

<tr align="center" bgcolor="#FCFCFC"><td width="10%" height='180' valign="top">
<table width="160" border="0" align="center" cellpadding="0" cellspacing="0"> 
<tr><td align="center" colspan="7">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr><td height="22" align="right">
<a href="chgrec_all_dtae.asp?date=<%= SubtractOneyear(dDate) %>"><&nbsp;</a>
<a href="chgrec_all_dtae.asp?date=<%= SubtractOneMonth(dDate) %>"><<</a></td> 
<td align="center"><font color="999999"><b><%= MonthName(Month(dDate)) & " " & Year(dDate) %></b></font></td>
<td><a href="chgrec_all_dtae.asp?date=<%= AddOneMonth(dDate) %>">>></a><a href="chgrec_all_dtae.asp?date=<%= AddOneYear(dDate) %>">&nbsp;></a>
</td></tr>
</table>
</td></tr> 
<tr> 
<td width="25" height="22" align="center"><font color="d08c00"><b>日</b></font></td> 
<td width="25" align="center"><b><font color="999999">一</font></b> </td> 
<td width="25" align="center"><b><font color="999999">二</font></b> </td> 
<td width="25" align="center"><b><font color="999999">三</font></b> </td> 
<td width="25" align="center"><b><font color="999999">四</font></b> </td> 
<td width="25" align="center"><b><font color="999999">五</font></b> </td> 
<td width="25" align="center"><b><font color="d08c00">六</font></b> </td> 
</tr> 
<% 
' 如果这个月的起始日期不是周日的话就加空的单元.' 
If iDOW <> 1 Then 
Response.Write vbTab & "<TR>" & vbCrLf 
iPosition = 1 
Do While iPosition < iDOW 
Response.Write vbTab & vbTab & "<TD></TD>" & vbCrLf 
iPosition = iPosition + 1 
Loop 
End If 

' 绘制这个月的日历 '
iCurrent = 1 
iPosition = iDOW 
Do While iCurrent <= iDIM 
' 如果是一行的开头就使用 TR 标记 '
If iPosition = 1 Then 
Response.Write vbTab & "<TR>" & vbCrLf 
End If 

' 如果这一天是我们选择的日期就高亮度显示该日期. '
If iCurrent = Day(dDate) Then 
Response.Write vbTab & vbTab & "<TD BGCOLOR=#eeeeee height=18 align=center><B>" & iCurrent & "</B></TD>" & vbCrLf 
Else 
Response.Write vbTab & vbTab & "<TD height=18 align=center><A HREF=""./chgrec_all_dtae.asp?date=" & Year(dDate) & "-" & Month(dDate)  & "-" & iCurrent & """>" & iCurrent & "</A></TD>" & vbCrLf 
End If 

' 如果满一周的话表格就另起一行' 
If iPosition = 7 Then 
Response.Write vbTab & "</TR>" & vbCrLf 
iPosition = 0 
End If 

iCurrent = iCurrent + 1 
iPosition = iPosition + 1 
Loop 

' 如果一个月不是以周六结束则加上相应的空单元.' 
If iPosition <> 1 Then 
Do While iPosition <= 7 
Response.Write vbTab & vbTab & "<TD> </TD>" & vbCrLf 
iPosition = iPosition + 1 
Loop 
Response.Write vbTab & "</TR>" & vbCrLf 
End If 
%> 
</table>

</td>
<td valign="top" height="180">
<table width="100%" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#c5b5e5">
<tr align="center" bgcolor="#FCFCFC">
<td width="4%" height ="20">No.</td>
<td width="7%">Line</td>
<td width="8%">机种</td>
<td width="8%">工程</td>
<td width="8%">工位</td>
<td width="8%">原作业者工号</td>
<td width="8%">姓名</td>
<td width="8%">原由</td>
<td width="8%">替位者工号</td>
<td width="8%">替位</td>
<td width="8%">职称</td>

<% 
mainbody()
CloseDB()
 %>
</table>

</td></tr>
</table>
</form>
</body>
</html>

<%


Sub MainBody()
If rs.eof then
  response.write "<tr bgcolor='#FCFCFC'><td width=""100%"" colspan=12  valign='top'>"
  response.write "没有记录。"
  response.write "</td></tr>"
End if

n=1 
While Not rs.eof and n<=rs.pagesize 
  response.write"<tr bgcolor='#FCFCFC' align='center'>"
   response.write "<td bgcolor='#FCFCFC' align ='center'>"&n&"</td>"
   response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("line")&"</td>"
    response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("机种")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("工程")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("工位")&"</td>" 
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("原作业者工号")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("姓名")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("原由")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("替位者工号")&"</td>"
      response.write "<td bgcolor='#e6F6F6' align ='center'>"&rs("替位")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("职称")&"</td>"
   response.write"</tr>"
   n=n+1 
  rs.movenext '显示页面的数据 '
Wend 
End Sub
%>


