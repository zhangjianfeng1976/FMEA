<!--#include file="../includes/db.asp"-->
<%
pgs =2500
select case request("cat")
    case ""
     sql="SELECT * FROM t_Fmea_Advice_history Where itmID = "& request("itmID")       
    case 2
     sql="SELECT * FROM t_Fmea_Advice_history Where Pjkey = '"&request("Pjkey")&"'"
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
<title>���ƴ�ʩ��Ŀһ��</title>
</head>

<body bgcolor ="#FCFCFC">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=13>
<b>���ƴ�ʩ�޶���¼</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="15%" rowspan ="2">����No.</td>
<td width="15%" rowspan ="2">�����ʩ</td>
<td width="5%"  rowspan ="2">������</td>
<td width="8%"  rowspan ="2">Ŀ�������</td>
<td width="18%" colspan ="5">��ʩ���</td>
</tr><tr bgcolor="#FCFCFC" align="center">
<td width="10%">���ô�ʩ</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RPN</td>
</tr>

<% 
call MainBody()
 %>

<tr>
<td width="75%" colspan=13 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">�ص���ҳ</a>][<a href="javascript:history.back()">����</a>]
</td><td  width="50%" align="right">
<% call clspage() %>   
</table>
</td>
</tr>
</table>
<% call closeDB() %>
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
    AdID = rs("AdID")
	ItmID = rs("ItmID")
	response.write"<tr bgcolor='#FCFCFC'>"	
	response.write"<td>"&rs("Pjkey")&"</td>"
    response.write"<td>"&rs("AdvContent")&"</td>"	
	response.write"<td>"&rs("Rspnser")&"</td>"
    response.write"<td>"&rs("FshDate")&"</td>"
    response.write"<td>"&rs("ActionContent")&"</td>"
    response.write"<td>"&rs("ResultS")&"</td>"
    response.write"<td>"&rs("ResultO")&"</td>"
    response.write"<td>"&rs("ResultD")&"</td>"
    response.write"<td>"&rs("ResultRPN")&"</td>" 
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=His_Advice_Browse.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=His_Advice_Browse.asp?page=1>��ҳ</a>&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&pagecount&">βҳ</a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub
%>

