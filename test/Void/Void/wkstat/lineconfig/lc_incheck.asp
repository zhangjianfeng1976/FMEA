<!--��Ա��֤-->

<!--#include file="../includes/db.asp"-->
<%
pgs =2500
select case request("cat")
       case "" 
       sql="SELECT top 10  * FROM  wt_employbase"
       pgs =25
       case 0
       sql="SELECT * FROM wt_employbase where cardid like '%"&request("Sc")&"%' ORDER BY cardid"                 
end select

call openDB()
rs.open sql,conn,1,1
if not rs.eof then
rs.pagesize =pgs
pagecount=rs.PageCount 
page=int(request.QueryString ("page"))
if page<=0 then page=1
if request.QueryString("page")="" then page=1
rs.AbsolutePage=page 
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>

<body>
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="20">����</td>
<td width="6%">����</td>
<td width="6%">ְ��</td>
<td width="6%">�Ա�</td>
<td width="4%">ϵ���</td>
<td width="25%">����</td>
<td width="8%">�볧����</td>
<td width="8%">��ְ��</td>
<td width="8%">״̬</td>
</tr>

<% call MainBody() %>
<% Call CloseDB() %>

</table>
</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=6>"
	response.write "û�м�¼��"
	response.write "</td></tr>"
end if

while not rs.eof and n<= rs.pagesize
	response.write"<tr bgcolor='#FCFCFC'>"
	response.write"<td>"&rs("CardID")&"</td>"	
	response.write"<td>"&rs("name")&"</td>"
        response.write"<td>"&rs("vocation")&"</td>"
        response.write"<td>"&rs("sex")&"</td>"
        response.write"<td>"&rs("DPcode")&"</td>"
        response.write"<td>"&rs("department")&"</td>"
	response.write"<td>"&rs("entrancedate")&"</td>"
        response.write"<td>"&rs("lastwd")&"</td>"
        response.write"<td>"&rs("statuss")&"</td>"
	response.write"</tr>"
	rs.MoveNext
wend
End Sub

%>
