<!--#include file="../includes/db.asp"-->
<%
LiID=request("liid")
if request("act") = "cancel" then
'response.write "<script language='JavaScript'>alert('��λ!');</script>" 
sta=request("sta")&"d"&year(now())&right("0"&month(now()),2)&right("0"&day(now()),2)
sqlx = "Update wt_lineconfig set wkstatus = 'N',statuspno ='"&sta&"'  where LineId = "&LiID
call openDB()
conn.execute(sqlx)
call closeDB()
Response.Redirect "lcbrowse_ld1.asp?cat=5&line="&request("line")&"&model="&request("model")

elseif request("act") = "allcancel" then

sta= "d"&year(now())&right("0"&month(now()),2)&right("0"&day(now()),2)
sqly = "Update wt_lineconfig set wkstatus = 'N',statuspno = rtrim(statuspno) + '"&trim(sta)&"'  where wkstatus = 'Y' and leader = '"&request.Cookies("leader")&"'"

call openDB()
conn.execute(sqly)
call closeDB()
'Response.Redirect "lcbrowse_ld1.asp"
end if

Select Case Request("OK") 
       Case "��ѯ" 	                
            select case Request("cata")
              case "�߼���ѯ"
              case "ģ����ѯ"                  
	            Response.Redirect "Lcbrowse_ld1.asp?Cat=0&Sc="&Request("Search")       		   	
              case "��Line��ѯ"                  
                    Response.Redirect "Lcbrowse_ld1.asp?Cat=1&Sc="&Request("Search")
              case "�����ֲ�ѯ"                 
                    Response.Redirect "Lcbrowse_ld1.asp?Cat=2&Sc="&Request("Search")
              case "�����̲�ѯ"                
                    Response.Redirect "Lcbrowse_ld1.asp?Cat=3&Sc="&Request("Search")
              case "����λ��ѯ"                 
                    Response.Redirect "Lcbrowse_ld1.asp?Cat=4&Sc="&Request("Search")              		
          end select
       Case "���й�λ��ѯ" 
             Response.Redirect "Lcbrowse.asp?Sc="&Request("Search")           
End Select

pgs =500
select case request("cat")
       case "" 
       sql="SELECT * FROM  v_wmaster where ��Ч ='Y' and ��� = '"&request.Cookies("leader")&"' order by LINE,����,����,��λ"
       pgs =2000
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
       case 5
         sql="SELECT * FROM v_wmaster where ��Ч ='Y' and ��� = '"&request.Cookies("leader")&"' and line = '"&request("line")&"' and ���� = '"&request("model")&"' order by line"
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

sql1="SELECT distinct LINE,���� FROM  v_wmaster where ��Ч ='Y' and ��� = '"&request.Cookies("leader")&"' order by LINE,����"
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sql1,conn,1,1


%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��λ��Ա����</title>
</head>

<body>
<form action="Lcbrowse_ld1.asp">
<table width="100%">
<tr bgcolor="#E6F7FF"><td height="30" align='center' bgcolor="#E6F7FF" colspan=11>

<table width="100%"><tr><td>
<b>[<%=request.Cookies("leader") %>]����Χ���̰���һ��</b>
[<a href=Lcbrowse_ld1.asp>��ʾȫ����λ</a>]</td>
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
<tr bgcolor="#E6F7FF"><td height="15" align='left' bgcolor="#E6F7FF" colspan=11>
����˵�����������LINE(���֣�,���Եõ����LINE(���֣��Ĺ������ã����в����õĹ�λ��ֱ�ӡ�<a href="Lcbrowse_ld1.asp?act=allcancel">��λ</a>����Ҫ�޶���Ա���á�ԭ��ҵ�߱��������Ҫ׷�ӹ�λ�������ϲ�ġ���λ��Ա���롿����EXCEL�������롣
</td></tr>
</td></tr>
<tr valign="top"><td width="15%">
<table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#FCFCFC" align="center">
<td width="15%" height="20">No.</td>
<td width="80%">Line(����)</td></tr>
<%
call mbd()
%>


</table>

</td><td  width="85%">

<table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="20">Line</td>
<td width="10%">����</td>
<td width="5%">����</td>
<td width="5%">��λ</td>
<td width="4%">����</td>
<td width="4%">����</td>
<td width="5%">ְ��</td>
<td width="10%">������</td>
<td width="20%" align="center">�������</td>
</tr>

<% 

call MainBody()

%>

<tr><td width="100%" colspan=11 height="30" bgcolor="#EFEFEF" align='center'>

<table width="100%">
<tr><td width="50%" align="left">
<% call clspage() %>
</td><td  width="50%" align="right">
[<a href="javascript:history.back()">����</a>]   
<% Call CloseDB() %>
</table>

</td></tr>

</table>
</form>
</body>
</html>

<%
Sub mbd()	
   if rs1.eof then
      response.write "<tr><td colspan=3>"
      response.write "��û�м�¼��"
      response.write "</td></tr>"
   else
      n=1
      while not rs1.eof
      response.write "<tr bgcolor='#FCFCFC'>"
      response.write "<td align ='left'>"&n&"</td>"
      response.write "<td align ='left'><a href ='lcbrowse_ld1.asp?cat=5&line="&rs1("line")&"&model="&rs1("����")&"'>"&rs1("line")&"("&rs1("����")&")</a></td>"
     ' response.write "<a href ='lcbrowse_ld1.asp?cat=5&line="&rs1("line")&"&model="&rs1("����")&"'>"&rs1("line")&"("&rs1("����")&")</a>&nbsp;&nbsp;"
      response.write "</tr>"
      rs1.MoveNext
      n=n+1
      wend
   end if

End Sub

Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=6>"
	response.write "û�м�¼��"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
        if rs("״̬")="δ��ְ" then
	response.write"<tr bgcolor='#FCFCFC'>"
        else
        response.write"<tr bgcolor='#FCACFC'>"
        end if	
        response.write"<td><a href=Lcbrowse_ld1.asp?cat=1&sc="&rs("Line")&">"&rs("Line")&"</a></td>"
        response.write"<td><a href=Lcbrowse_ld1.asp?cat=2&sc="&rs("����")&">"&rs("����")&"</a></td>"
        response.write"<td><a href=Lcbrowse_ld1.asp?cat=3&sc="&rs("����")&">"&rs("����")&"</a></td>"
        response.write"<td><a href=Lcbrowse_ld1.asp?cat=4&sc="&rs("��λ")&">"&rs("��λ")&"</a></td>"
	response.write"<td><a href=../wokermanage/emrecord.asp?CardID="&rs("����")&">"&rs("����")&"</a></td>"
        response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("ְ��")&"</td>"
        response.write"<td>"&rs("������")&"</td>"
	response.write"<td>&nbsp;&nbsp;"
        response.write"<a href=Lcbrowse_ld1.asp?act=cancel&liid="&rs("lineid")&"&line="&trim(rs("line"))&"&model="&trim(rs("����"))&"&sta="&trim(rs("�汾"))&">��λ</a>&nbsp;&nbsp;"
        response.write"<a href=LcEdit.asp?lineID="&rs("lineid")&">ԭ��ҵ�߱��</a>&nbsp;&nbsp;"
        if hour(now()) < 10 then   
        response.write"<a href=LcAdd.asp?lineID="&rs("lineid")&">��λ����</a>&nbsp;&nbsp;"
        else
        response.write"�ϰ�,��10����"
        end if
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=Lcbrowse_ld1.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld1.asp?page="&pagecount&">βҳ<a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=Lcbrowse_ld1.asp?page=1>��ҳ<a>&nbsp;&nbsp;"
Response.Write "<a href=Lcbrowse_ld1.asp?page="&page-1&">��ҳ<a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=Lcbrowse_ld1.asp?page=1>��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld1.asp?page="&page-1&">��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld1.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld1.asp?page="&pagecount&">βҳ<a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����λ"
end sub


%>