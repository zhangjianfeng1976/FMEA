<!--#include file="../includes/db.asp"-->
<%
 sql="SELECT * FROM  v_drrate order by �Ʊ�"
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
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=14>
<b>ֱ�����Ա�����б�</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="10%">�Ʊ�</td>
<td width="4%" bgcolor=#e6F6F6>��Ա����</td>
<td width="4%">���1��</td>
<td width="4%">���2��</td>
<td width="4%">�鳤1��</td>
<td width="4%">�鳤2��</td>
<td width="4%" bgcolor=#e6F6F6>�����Ա</td>
<td width="4%">MEISTER B</td>
<td width="4%">MEISTER C</td>
<td width="4%">Ա��2��</td>
<td width="4%">Ա��3��</td>
<td width="4%" bgcolor=#e6F6F6>ֱ����Ա</td>
<td width="4%">�鳤Ա����</td>
<td width="4%">ֱ���</td>
</tr>

<% call MainBody() %>


</table>
<% Call CloseDB() %>
</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=14>"
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
        response.write"<td>"&depart&"</td>"
        response.write"<td bgcolor=#e6F6F6>"&rs("��Ա��")&"</td>"
        a = a +	rs("��Ա��")
	response.write"<td>"&rs("���1��")&"</td>"
        b= b + rs("���1��")
        response.write"<td>"&rs("���2��")&"</td>"
        c= c + rs("���2��")
        response.write"<td>"&rs("�鳤1��")&"</td>"
        d= d + rs("�鳤1��")
        response.write"<td>"&rs("�鳤2��")&"</td>"
        e= e + rs("�鳤2��")
        response.write"<td bgcolor=#e6F6F6>"&rs("�����Ա")&"</td>"
        f= f + rs("�����Ա")
        response.write"<td>"&rs("MEISTER_B")&"</td>"
        g= g + rs("MEISTER_B")
        response.write"<td>"&rs("MEISTER_C")&"</td>"
        h= h + rs("MEISTER_C")
        response.write"<td>"&rs("Ա��2��")&"</td>"
        i= i + rs("Ա��2��")
        response.write"<td>"&rs("Ա��3��")&"</td>"
        j= j + rs("Ա��3��")
        response.write"<td bgcolor=#e6F6F6>"&rs("ֱ����Ա")&"</td>"
        k= k + rs("ֱ����Ա")

        x = rs("�鳤Ա����")

     '   if x <> 0 then
     '    y = 100/x
     '   else 
     '    y = x
     '   end if

        response.write"<td>"&x&"%</td>"
        xdirect = rs("�����Ա")
        if xdirect = 0 then
        response.write "<td>None</td>"
        else
        response.write"<td>"&round(rs("ֱ����Ա")/xdirect,2)&"</td>"
	end if
        response.write"</tr>"
	rs.MoveNext
n=n+1
wend
       response.write"<tr bgcolor='#F59595'>"
       response.write"<td>�ϼ�</td>"
       response.write"<td>"&a&"</td>"	
       response.write"<td>"&b&"</td>"
       response.write"<td>"&c&"</td>"
       response.write"<td>"&d&"</td>"
       response.write"<td>"&e&"</td>"
       response.write"<td>"&f&"</td>"
       response.write"<td>"&g&"</td>"
       response.write"<td>"&h&"</td>"
       response.write"<td>"&i&"</td>"
       response.write"<td>"&j&"</td>"
       response.write"<td>"&k&"</td>"
       response.write"<td>"&round((g+ h + i + j)/(d + e),2)&"<div>(Ա���鳤��)</div></td>"
       response.write"<td>"&round(k/f,2)&"</td>"
       response.write"</tr>"
End Sub
%>
