<!--#Include file="includes/db.asp"-->
<% 
' Ҫ���õĺ������� '
'������ݼ��·ݵõ�ÿ�µ������� '
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
'�õ�һ���¿�ʼ������. '
Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth) 
Dim dTemp 
dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth) 
GetWeekdayMonthStartsOn = WeekDay(dTemp) 
End Function 
'�õ���ǰһ���µ���һ����. '
Function SubtractOneMonth(dDate) 
SubtractOneMonth = DateAdd("m", -1, dDate) 
End Function 
'�õ���ǰһ���µ���һ����. '
Function AddOneMonth(dDate) 
AddOneMonth = DateAdd("m", 1, dDate) 
End Function 
'�õ���ǰһ�����һ��. '
Function SubtractOneYear(dDate) 
SubtractOneYear = DateAdd("yyyy", -1, dDate) 
End Function 
'�õ���ǰһ�����һ��.' 
Function AddOneYear(dDate) 
AddOneYear = DateAdd("yyyy", 1, dDate) 
End Function 
' ������������' 

Dim dDate ' ������ʾ������ '
Dim iDOW ' ÿһ�¿�ʼ������ '
Dim iCurrent ' ��ǰ���� '
Dim iPosition ' ����еĵ�ǰλ�� '


' �õ�ѡ������ڲ�������ڵĺϷ��� '
If IsDate(Request("date")) Then 
dDate = CDate(Request("date")) 
Else 
If IsDate(Request("year") & "-" & Request("month") & "-" & Request("day")) Then 
dDate = CDate(Request("year") & "-" & Request("month") & "-" & Request("day")) 
Else 
dDate = Date() 

If Len(Request("month")) <> 0 Or Len(Request("day")) <> 0 Or Len(Request("year")) <> 0 Or Len(Request("date")) <> 0 Then 
Response.Write "����ѡ������ڸ�ʽ����ȷ��ϵͳ��ʹ�õ�ǰ����.<BR><BR>" 
End If 

End If 
End If 

'�õ����ں������ȵõ�����µ�����������µ���ʼ����. '
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

sql="SELECT  * FROM v_wdairy where ���� = '"&aa&"'"

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
<title>��Ա��λ����</title>
<link href="Includes/Style.css" rel="stylesheet" type="text/css">
</head>

<body>	
<form action="chgrec_all_dtae.asp">
<table width="100%" border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr><td height="35" align='center' bgcolor="#E6F7FF" colspan='3'>
<b>������λһ����[<%= aa %>]</b>
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
<td width="25" height="22" align="center"><font color="d08c00"><b>��</b></font></td> 
<td width="25" align="center"><b><font color="999999">һ</font></b> </td> 
<td width="25" align="center"><b><font color="999999">��</font></b> </td> 
<td width="25" align="center"><b><font color="999999">��</font></b> </td> 
<td width="25" align="center"><b><font color="999999">��</font></b> </td> 
<td width="25" align="center"><b><font color="999999">��</font></b> </td> 
<td width="25" align="center"><b><font color="d08c00">��</font></b> </td> 
</tr> 
<% 
' �������µ���ʼ���ڲ������յĻ��ͼӿյĵ�Ԫ.' 
If iDOW <> 1 Then 
Response.Write vbTab & "<TR>" & vbCrLf 
iPosition = 1 
Do While iPosition < iDOW 
Response.Write vbTab & vbTab & "<TD></TD>" & vbCrLf 
iPosition = iPosition + 1 
Loop 
End If 

' ��������µ����� '
iCurrent = 1 
iPosition = iDOW 
Do While iCurrent <= iDIM 
' �����һ�еĿ�ͷ��ʹ�� TR ��� '
If iPosition = 1 Then 
Response.Write vbTab & "<TR>" & vbCrLf 
End If 

' �����һ��������ѡ������ھ͸�������ʾ������. '
If iCurrent = Day(dDate) Then 
Response.Write vbTab & vbTab & "<TD BGCOLOR=#eeeeee height=18 align=center><B>" & iCurrent & "</B></TD>" & vbCrLf 
Else 
Response.Write vbTab & vbTab & "<TD height=18 align=center><A HREF=""./chgrec_all_dtae.asp?date=" & Year(dDate) & "-" & Month(dDate)  & "-" & iCurrent & """>" & iCurrent & "</A></TD>" & vbCrLf 
End If 

' �����һ�ܵĻ���������һ��' 
If iPosition = 7 Then 
Response.Write vbTab & "</TR>" & vbCrLf 
iPosition = 0 
End If 

iCurrent = iCurrent + 1 
iPosition = iPosition + 1 
Loop 

' ���һ���²��������������������Ӧ�Ŀյ�Ԫ.' 
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
<td width="8%">����</td>
<td width="8%">����</td>
<td width="8%">��λ</td>
<td width="8%">ԭ��ҵ�߹���</td>
<td width="8%">����</td>
<td width="8%">ԭ��</td>
<td width="8%">��λ�߹���</td>
<td width="8%">��λ</td>
<td width="8%">ְ��</td>

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
  response.write "û�м�¼��"
  response.write "</td></tr>"
End if

n=1 
While Not rs.eof and n<=rs.pagesize 
  response.write"<tr bgcolor='#FCFCFC' align='center'>"
   response.write "<td bgcolor='#FCFCFC' align ='center'>"&n&"</td>"
   response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("line")&"</td>"
    response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("����")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("����")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("��λ")&"</td>" 
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("ԭ��ҵ�߹���")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("����")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("ԭ��")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("��λ�߹���")&"</td>"
      response.write "<td bgcolor='#e6F6F6' align ='center'>"&rs("��λ")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("ְ��")&"</td>"
   response.write"</tr>"
   n=n+1 
  rs.movenext '��ʾҳ������� '
Wend 
End Sub
%>


