<!--#include file="../includes/db.asp"-->
<%
pgs =2500
select case request("cat")
    case ""
        sql="SELECT * FROM t_Fmea_Item_history where  ItmID = "&request("ItmID")
    case 2
        sql="SELECT * FROM t_Fmea_Item_history where  Pjkey = '"&request("Pjkey")&"' order by subpjname,pageno,workqueue"
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
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE2 {color: #0000FF}
-->
</style>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ֹ�����Ŀ�޶�һ��</title>
</head>

<body bgcolor ="#FCFCFC">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=18>
<b><% = request("Pjkey") %>��ҵ���������޶�һ��</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="2%">���</td>
<td width="2%">˳��</td>
<td width="5%" height="20">��������</td>
<td width="5%">��׼��ҳ��</td>
<td width="7%">��Ʒ���</td>
<td width="10%">��Ʒ���Ƽ�����</td>
<td width="8%">��ҵ����������</td>
<td width="10%">Ǳ��ȱ��ģʽ</td>
<td width="10%">Ǳ��ȱ�ݺ��</td>
<td width="7%">���й�����Ƽ����̿���</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RPN</td>
<td width="2%">��ʩ</td> 
</tr>

<% 
call MainBody() 
%>

<tr>
<td width="75%" colspan=18 height="30" bgcolor="#EFEFEF" align='center'>
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
	ItmID = rs("ItmID")
	response.write"<tr bgcolor='#FCFCFC'>"
	response.write"<td>"&rs("Pjkey")&"</td>"
    response.write"<td>"&rs("WorkQueue")&"</td>"
    response.write"<td>"&rs("SubPjName")&"</td>"
    response.write"<td>"&rs("PageNo")&"</td>"
	response.write"<td>"&rs("PartsNo")&"</a></td>"
    response.write"<td><div>("&rs("PartsType")&")</div>"&rs("PartsName")&"</td>"
    response.write"<td><div>("&rs("WorkCata")&")</div><div>"&rs("WorkContent")&"</div></td>"
    response.write"<td>"&rs("NGMode")&"</td>"
	response.write"<td>"&rs("NGEffect")&"</td>"
    response.write"<td>"&rs("ProDesign")&"</td>"
    response.write"<td align ='center'>"&rs("Sev")&"</td>"
    response.write"<td align ='center'>"&rs("Occ")&"</td>"
    response.write"<td align ='center'>"&rs("Det")&"</td>"
    
    vRPN = trim(rs("RPN"))
    if vRPN > 2 Then
        response.write"<td bgcolor = '#DDFFAA'><span class='STYLE1'>&nbsp;"& vRPN &"</span></td>"    
        response.write"<td bgcolor = '#DDFFAA'><a href=his_Advice_Browse.asp?ItmID="&ItmID&">��Ӧ</a></td>" 
        response.write"</tr>"
    else
        response.write"<td>"& vRPN &"</td>" 
        response.write"<td></td>"
        response.write"</tr>"
    end if
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=His_Process_Browse.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=His_Process_Browse.asp?page=1>��ҳ</a>&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&pagecount&">βҳ</a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub
%>

