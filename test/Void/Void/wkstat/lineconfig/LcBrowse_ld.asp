<!--#include file="../includes/db.asp"-->
<%
LiID=request("liid")
if request("act") = "cancel" then
'response.write "<script language='JavaScript'>alert('��λ!');</script>" 
sqlx = "Update wt_lineconfig set wkstatus = 'N' where LineId = "&LiID
call openDB()
conn.execute(sqlx)
call closeDB()
end if

Select Case Request("OK") 
       Case "��ѯ" 	                
            select case Request("cata")
              case "�߼���ѯ"
              case "ģ����ѯ"                  
	            Response.Redirect "Lcbrowse_ld.asp?Cat=0&Sc="&Request("Search")       		   	
              case "��Line��ѯ"                  
                    Response.Redirect "Lcbrowse_ld.asp?Cat=1&Sc="&Request("Search")
              case "�����ֲ�ѯ"                 
                    Response.Redirect "Lcbrowse_ld.asp?Cat=2&Sc="&Request("Search")
              case "�����̲�ѯ"                
                    Response.Redirect "Lcbrowse_ld.asp?Cat=3&Sc="&Request("Search")
              case "����λ��ѯ"                 
                    Response.Redirect "Lcbrowse_ld.asp?Cat=4&Sc="&Request("Search")              		
          end select
       Case "���й�λ��ѯ" 
             Response.Redirect "Lcbrowse.asp?Sc="&Request("Search")           
End Select

pgs =500
select case request("cat")
       case "" 
       sql="SELECT * FROM  v_wmaster where ��Ч ='Y' and ��� = '"&request.Cookies("leader")&"' order by LINE,����,����,��λ"
       pgs =200
       case 0
       sql="SELECT * FROM v_wmaster where ��Ч ='Y' and ��� = '"&request.Cookies("leader")&"' and (line like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' Or ��λ like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%') order by LINE,����,����,��λ"     
       case 1
         sql="SELECT * FROM v_wmaster where ��Ч ='Y' and ��� = '"&request.Cookies("leader")&"' and line = '"&request("Sc")&"' order by ����"
       case 2
         sql="SELECT * FROM v_wmaster where ��Ч ='Y' and ��� = '"&request.Cookies("leader")&"' and ���� = '"&request("Sc")&"' order by line"
       case 3
         sql="SELECT * FROM v_wmaster where ��Ч ='Y' and ��� = '"&request.Cookies("leader")&"' and ���� = '"&request("Sc")&"' order by line"  
       case 4
         sql="SELECT * FROM v_wmaster where ��Ч ='Y' and ��� = '"&request.Cookies("leader")&"' and ��λ = '"&request("Sc")&"' order by line" 
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
<form action="Lcbrowse_ld.asp">
<table  width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=11>
<table width="100%"><tr><td>
<b>[<%=request.Cookies("leader") %>]����Χ���̰���һ��</b>
[<a href=lcbrowse_ld.asp>��ʾȫ����λ</a>]</td>
<td align='right' bgcolor="#e6f7ff" colspan ='10'>
<select name='cata' class="smallinput">
<option>ģ����ѯ</option>
<option>��Line��ѯ</option>
<option>�����ֲ�ѯ</option>
<option>�����̲�ѯ</option>
<option>����λ��ѯ</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="��ѯ" name="OK" class="smallInput">
<input type="submit" value="���й�λ��ѯ" name="OK" class="smallInput"> 
��<a href="../includes/LineConfigBatch.xls">��λ��Ա����</a>��
</td></tr>

</table>
</td></tr>

<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="25">Line</td>
<td width="15%">����</td>
<td width="5%">����</td>
<td width="8%">��λ</td>
<td width="5%">����</td>
<td width="5%">����</td>
<td width="5%">ְ��</td>
<td width="8%">������</td>
<td width="20%" align="center">�������</td>
</tr>

<% 

call MainBody()

%>

<tr>
<td width="100%" colspan=11 height="30" bgcolor="#EFEFEF" align='center'>
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
	response.write "<tr><td colspan=6>"
	response.write "û�м�¼��"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
	response.write"<tr bgcolor='#FCFCFC'>"	
        response.write"<td><a href=lcbrowse_ld.asp?cat=1&sc="&rs("Line")&">"&rs("Line")&"</a></td>"
        response.write"<td><a href=lcbrowse_ld.asp?cat=2&sc="&rs("����")&">"&rs("����")&"</a></td>"
        response.write"<td><a href=lcbrowse_ld.asp?cat=3&sc="&rs("����")&">"&rs("����")&"</a></td>"
        response.write"<td><a href=lcbrowse_ld.asp?cat=4&sc="&rs("��λ")&">"&rs("��λ")&"</a></td>"
	response.write"<td><a href=../wokermanage/emrecord.asp?CardID="&rs("����")&">"&rs("����")&"</a></td>"
        response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("ְ��")&"</td>"
        response.write"<td>"&rs("������")&"</td>"
	response.write"<td>&nbsp;&nbsp;"
       ' response.write"<a href=lcbrowse_ld.asp?act=cancel&liid="&rs("lineid")&">��λ</a>&nbsp;&nbsp;"
       ' response.write"<a href=LcEdit.asp?lineID="&rs("lineid")&">ԭ��ҵ�߱��</a>&nbsp;&nbsp;"
       ' if hour(now()) < 10 then   
       ' response.write"<a href=LcAdd.asp?lineID="&rs("lineid")&">��λ����</a>&nbsp;&nbsp;"
      '  else
       ' response.write"�ϰ�,��10����"
       ' end if
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&pagecount&">βҳ<a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=Lcbrowse_ld.asp?page=1>��ҳ<a>&nbsp;&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&page-1&">��ҳ<a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=Lcbrowse_ld.asp?page=1>��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&page-1&">��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&pagecount&">βҳ<a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����λ"
end sub


%>