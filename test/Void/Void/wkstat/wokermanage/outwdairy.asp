<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "��ѯ" 	                
            select case Request("cata")
              case "�߼���ѯ"
              case "ģ����ѯ"
	           Response.Redirect "outwdairy.asp?Cat=0&Sc="&Request("Search")			   		
              case "��������Ա"
                    Response.Redirect "outwdairy.asp?Cat=1&Sc="&Request("Search")
          end select            
End Select

pgs =2500
select case request("cat")
       case "" 
       sql="SELECT * FROM  v_outwdairy order by ���� DESC"
       pgs =25
       case 0
       sql="SELECT * FROM v_outwdairy where ���� like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' Or ְ�� like '%"&request("Sc")& _
           "%' Or ϵ��� like '%"&request("Sc")& _
           "%' Or ԭ�� like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' ORDER BY ����"    
       case 1
                  
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
<title>��λ��Ա����</title>
</head>

<body>
<form action="outwdairy.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr><td height="20" align='right' bgcolor="#e6f7ff">
<select name='cata' class="smallinput">
<option>ģ����ѯ</option>

</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="��ѯ" name="OK" class="smallInput">
</td>
</tr>
</table>
</form>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=10>
<b>��������Ա����״����</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="20">����</td>
<td width="5%">�·�</td>
<td width="8%">����</td>
<td width="4%">����</td>
<td width="4%">ְ��</td>
<td width="8%">ԭ��</td>
<td width="25%">����</td>
<td width="8%">ϵ���</td>
</tr>

<% call MainBody() %>

<tr>
<td width="75%" colspan=10 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
<% call clspage() %>
</td><td  width="50%" align="right">
[<a href="javascript:history.back()">����</a>]   
<% Call CloseDB() %>
</table>
</td>
</tr>
</table>
</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=10>"
	response.write "û�м�¼��"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize

	response.write"<tr bgcolor='#FCFCFC'>"
	response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("�·�")&"</td>"
        response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("ְ��")&"</td>"
	response.write"<td>"&rs("ԭ��")&"</td>"
        response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("ϵ���")&"</td>"
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&pagecount&">βҳ<a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=outwdairy.asp?page=1>��ҳ<a>&nbsp;&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&page-1&">��ҳ<a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=outwdairy.asp?page=1>��ҳ<a>&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&page-1&">��ҳ<a>&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&pagecount&">βҳ<a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub


%>
