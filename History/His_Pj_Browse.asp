<!--#include file="../includes/db.asp"-->
<%
sql = "SELECT * FROM t_Fmea_PJ_history where PjID ="&request("PjID")
pgs = 25

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
<title>���ֹ���</title>
</head>

<body bgcolor ="#FCFCFC">
<form action="His_Pj_Browse.asp">

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=11>
<b>�汾�޶���¼�б�</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="20">FMEA��ż��汾</td>
<td width="10%">��Ŀ</td>
<td width="12%">����</td>
<td width="8%">����</td>
<td width="5%">������</td>
<td width="10%">��������</td>
<td width="10%">��Ҫ������</td>
<td width="8%">FMEA��ʼ����</td>
<td width="8%">��������</td>
</tr>

<% 
  call MainBody()
%>

<tr>
<td width="75%" colspan=10 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">�ص���ҳ</a>] [<a href="javascript:history.back()">����</a>]  
</td><td  width="50%" align="right">
<% call clspage() %> 
</table>
</td>
</tr>
</table>
<% call closeDB() %>

</form>

</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=6>"
	response.write "û�м�¼��"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize

	PjKey = rs("PjKey")
    Serial = rs("Serial")
    
	response.write"<tr bgcolor='#FCFCFC'>"
	response.write"<td><div>"&Pjkey&"</div>"
    response.write"<div>(<a href=His_Process_Browse.asp?cat=2&Pjkey="&Pjkey&">�����޶�</a>)"
    response.write"(<a href=His_Advice_Browse.asp?cat=2&Pjkey="&Pjkey&">��Ӧ�޶�</a>)"
    response.write"</div></td>"
	response.write"<td>"&rs("ProjectName")&"</td>"
    response.write"<td>"&rs("ItemName")&"</td>"
    response.write"<td>"&rs("Model")&"</td>"
    response.write"<td>"&rs("Author")&"</td>"
    response.write"<td>"&rs("BurdenDept")&"</td>"
	response.write"<td>"&rs("Onduty")&"</td>"
    response.write"<td>"&rs("CreateDate")&"</td>"
    response.write"<td>"&rs("ModifyDate")&"</td>"
	response.write"</td></tr>"
	
	rs.MoveNext
	
n=n+1
wend

End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=His_Pj_Browse.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=His_Pj_Browse.asp?page=1>��ҳ</a>&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&pagecount&">βҳ</a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub

%>

