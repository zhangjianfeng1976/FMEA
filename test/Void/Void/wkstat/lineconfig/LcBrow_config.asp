<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "��ѯ" 	                
            select case Request("cata")
              case "�߼���ѯ"
              case "ģ����ѯ"
	            Response.Redirect "LcBrow_config.asp?Cat=0&Sc="&Request("Search")			   		
              case "��Ч����"
                    Response.Redirect "LcBrow_config.asp?Cat=1&Sc="&Request("Search")
              case "���䵣��"
                    Response.Redirect "LcBrow_config.asp?Cat=2&Sc="&Request("Search")
              case "��������ѯ"
                    Response.Redirect "LcBrow_config.asp?Cat=3&Sc="&Request("Search")
              case "��������ѯ"
                    Response.Redirect "LcBrow_config.asp?Cat=4&Sc="&Request("Search")		
          end select            
End Select


lineID=request("lineID")
Select case request("action")
       case "changeN"
            sql = "Update wt_lineconfig set wkstatus = 'N' where LineId = "&LineID
            call openDB()
            conn.execute(sql)
            call closeDB()
       case "changeY"
            sql = "Update wt_lineconfig set wkstatus = 'Y' where LineId = "&LineID
            call openDB()
            conn.execute(sql)
            call closeDB()
end select

select case request("cat")
       case "" 
       sql="SELECT * FROM  wt_lineconfig"
       pgs =25
       case 0
       sql="SELECT * FROM wt_lineconfig where lineid like '%"&request("Sc")& _
           "%' Or line like '%"&request("Sc")& _
           "%' Or model like '%"&request("Sc")& _
           "%' Or wogroup like '%"&request("Sc")& _
           "%' Or workstation like '%"&request("Sc")& _
           "%' Or originworker like '%"&request("Sc")& _
           "%' ORDER BY lineid" 
       pgs =2500      
       case 1
       sql="SELECT * FROM  wt_lineconfig where Wkstatus ='N'"
       pgs =2500
       case 2
       sql="SELECT * FROM  wt_lineconfig where changedate is not null"
       pgs =2500       
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
<form action="LcBrow_config.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr><td height="20" align='right' bgcolor="#e6f7ff">
<select name='cata' class="smallinput">
<option>ģ����ѯ</option>
<option>��Ч����</option>
<option>���䵣��</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="��ѯ" name="OK" class="smallInput">
</td>
</tr>
</table>
</form>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=9>
<b>��������</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="left">
<td width="100%" height="20" colspan = 9>
*�¼ǡ���Ч����ָ���Ǽ�¼��״̬�Ƿ���Ч�������Y�Ļ��������Ա���Ե����˹�λ�������N�Ļ�����ʾ�ü�¼����Ч����λ�ѱ���һԱ�����
</td></tr>
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="20">Line</td>
<td width="8%">����</td>
<td width="10%">����</td>
<td width="10%">��λ</td>
<td width="10%">ԭ������</td>
<td width="4%">��Ч</td>
<td width="8%">��������</td>
<!--<td width="20%">����</td>-->
</tr>

<% call MainBody() %>

<tr>
<td width="75%" colspan=9 height="30" bgcolor="#EFEFEF" align='center'>
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
	response.write "<tr><td colspan=6>"
	response.write "û�м�¼��"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
	response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<td>"&rs("Line")&"</td>"	
        response.write"<td>"&rs("Model")&"</td>"
        response.write"<td>"&rs("WoGroup")&"</td>"
        response.write"<td>"&rs("WorkStation")&"</td>"
	response.write"<td>"&rs("Originworker")&"</td>"
        response.write"<td>"&rs("Wkstatus")&"</td>"
        response.write"<td>"&rs("changedate")&"</td>"
'       response.write"<td><a href=LcEdit.asp?lineID="&rs("lineid")&">���ԭ������</a>&nbsp;&nbsp;"
'       response.write"<a href=LcBrow_config.asp?action=changeN&lineID="&rs("lineid")&">תΪ��Ч</a>&nbsp;&nbsp;"
'       response.write"<a href=LcBrow_config.asp?action=changeY&lineID="&rs("lineid")&">תΪ��Ч</a>&nbsp;&nbsp;"
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=LcBrow_config.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=LcBrow_config.asp?page=1>��ҳ<a>&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&page-1&">��ҳ</a>&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&pagecount&">βҳ</a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub


%>
