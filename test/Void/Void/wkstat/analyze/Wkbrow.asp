<!--#include file="../includes/db.asp"-->
<%
select case request("cat")
  case ""
  sql="SELECT * FROM  v_linestruture order by �Ʊ�,ְ��"
  case 1
  sql="SELECT * FROM  v_linestruture where �Ʊ� = '"&request("ka")&"' order by �Ʊ�,ְ��"
  case 2
  sql="SELECT * FROM  v_linestruture where ְ�� like '%"&request("ka")&"%' order by �Ʊ�,ְ��"
end select

call openDB()
rs.open sql,conn,1,1
if not rs.eof then
rs.pagesize = 100
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
<title>��Ա����</title>
</head>

<body>
<table width="70%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=9>
<b>��Ա�����б�</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="10%">�Ʊ�</td>
<td width="10%">ְλ�ֲ�</td>
<td width="4%">��ϯ</td>
<td width="4%">������</td>
<td width="4%">������</td>
<td width="4%">������</td>
<td width="4%">���������</td>
</tr>

<% call MainBody() %>


</table>
<% Call CloseDB() %>
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
      select case rs("�Ʊ�")
            case "A41"
              depart ="����1���첿����1��"
            case "A42"
              depart ="����1���첿����2��"
            case "A48"
              depart ="MFP1���첿�����ƽ���"
            case "A54"
              depart ="����2���첿����1��"
            case "A55"
              depart ="����2���첿����2��"
            case "A61"
              depart ="����3���첿����1��"
            case "A62"
              depart ="����3���첿����2��"
            case "A91"
              depart = "�����ƽ���"
            case else
              depart ="δ����"
      end select

	response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<td><a href=wkbrow.asp?cat=1&ka="&rs("�Ʊ�")&">"&depart&"</a></td>"	
	response.write"<td><a href=wkbrow.asp?cat=2&ka="&rs("ְ��")&">"&rs("ְ��")&"</a></td>"
        response.write"<td>"&rs("��ϯ")&"</td>"
        i0 = rs("��ϯ")
        a=a+ i0
        i1 = rs("����")
        if isnull(i1) then 
           i1 = 0
        end if
        response.write"<td>"&i1&"</td>"
        b=b+ i1
        i2 = rs("������")
        if isnull(i2) then 
           i2 = 0
        end if
        response.write"<td>"&i2&"</td>" 
        c=c+ i2
        response.write"<td>"&i0 - i1  -i2 &"</td>"
        response.write"<td>"&round((i1+i2)/i0*100,2)&"%</td>"
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend

        response.write"<tr  bgcolor='#F59595'>"
        response.write"<td>[<a href='javascript:history.back()'>����</a>] </td>"	
	response.write"<td>�ϼ�</td>"      
        response.write"<td>"&a&"</td>"
        response.write"<td>"&b&"</td>"
        response.write"<td>"&c&"</td>"
        response.write"<td>"&a-b-c&"</td>"
        response.write"<td>"&round((b+c)/a*100,2)&"%</td>"
	response.write"</tr>"
End Sub
%>
