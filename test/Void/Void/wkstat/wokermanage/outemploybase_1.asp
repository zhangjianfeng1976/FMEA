<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "��ѯ" 	                
            select case Request("cata")
              case "�߼���ѯ"
              case "ģ����ѯ"
	    Response.Redirect "outemploybase_1.asp?Cat=0&Sc="&Request("Search")			   		
              case "��������Ա"
                    Response.Redirect "outemploybase_1.asp?Cat=1&Sc="&Request("Search")
          end select            
     Case "����"
       Response.Redirect "add.asp"
End Select


pgs =2500
select case request("cat")
       case "" 
       sql="SELECT * FROM  v_outemploybase_1"
       pgs =100
       case 0
       sql="SELECT * FROM v_outemploybase_1 where cardid like '%"&request("Sc")& _
           "%' Or name like '%"&request("Sc")& _
           "%' Or vocation like '%"&request("Sc")& _
           "%' Or dpcode like '%"&request("Sc")& _
           "%' Or department like '%"&request("Sc")& _
           "%' Or sex like '%"&request("Sc")& _
           "%' ORDER BY cardid"    
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
<form action="outemploybase_1.asp">


<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=10>
<b>��������Ա�б�</b></td>
</tr>

<tr>
<td height="20" align='left' bgcolor="#e6f7ff" colspan="6">
���ȷ�ϣ�
<a href="outemploybase_1.asp?Cat=0&Sc=A4">����1���첿</a>&nbsp;&nbsp;
<a href="outemploybase_1.asp?Cat=0&Sc=A5">����2���첿</a>&nbsp;&nbsp;
<a href="outemploybase_1.asp?Cat=0&Sc=A6">����3���첿</a>&nbsp;&nbsp;
</td>
<td align='right' bgcolor="#e6f7ff" colspan ="4">
<select name='cata' class="smallinput">
<option>ģ����ѯ</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="��ѯ" name="OK" class="smallInput">
<%
'  response.write "<input type='submit' value='����' name='OK' class='smallInput'>"
%>
</td>
</tr>

<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="20">����</td>
<td width="5%">����</td>
<td width="8%">ְ��</td>
<td width="4%">�Ա�</td>
<td width="4%">ϵ���</td>
<td width="30%">����</td>
<td width="8%">�볧����</td>
<td width="8%">��ְ��</td>
<td width="8%">״̬</td>
<td width="8%" align="center">��������Ա����</td> 
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
</form>
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
	CardID=rs("CardID")
	response.write"<tr bgcolor='#FCFCFC'>"
	response.write"<td><a href=emrecord.asp?CardID="&rs("CardID")&">" &CardID&"</a></td>"	
	response.write"<td>"&rs("name")&"</td>"
        response.write"<td>"&rs("vocation")&"</td>"
        response.write"<td>"&rs("sex")&"</td>"
        response.write"<td>"&rs("dpcode")&"</td>"
        response.write"<td>"&rs("department")&"</td>"
	response.write"<td>"&rs("entrancedate")&"</td>"
        response.write"<td>"&rs("lastwd")&"</td>"
        response.write"<td>"&rs("statuss")&"</td>"
 	response.write"<td><a href=emEdit.asp?ID="&rs("id")&">"&"������</a>&nbsp;&nbsp;&nbsp;&nbsp;"
       ' response.write"<a href=add.asp?CardID="&rs("CardID")&">"&"������λ</a>&nbsp;&nbsp;"
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=outemploybase_1.asp?page=1>��ҳ<a>&nbsp;&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=outemploybase_1.asp?page=1>��ҳ</a>&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&page-1&">��ҳ</a>&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&pagecount&">βҳ</a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub


%>
