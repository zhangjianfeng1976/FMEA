<!--��LINE��λ��-->
<!--#include file="../includes/db.asp"-->
<%
LiID=request("liid")
if request("act") = "cancel" then
sqlx = "Update wt_lineconfig set wkstatus = 'N' where LineId = "&LiID
call openDB()
conn.execute(sqlx)
call closeDB()
end if


sql="SELECT  * FROM v_lc_count"
sql1="SELECT * FROM  v_wmaster where ��Ч ='Y' and line = '"&request("line")&"' and ���� = '"&request("model")&"'"
Call OpenDB
rs.open sql,conn,1,1
Set rs_detail = Server.CreateObject("ADODB.Recordset")
rs_detail.open sql1,conn,1,1


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Includes/Style.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<body>
<table width="100%" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr bgcolor="#E6F7FF">
<td  colspan="2">
<b>��LINE/���ֹ������ñ�</b> 
</td></tr> 
<tr><td width="300" valign="top">
<iframe src='lc_count_in.asp' name='container1' width='100%' marginwidth='1' height='470' marginheight='1' frameborder='0' scrolling='auto'></iframe>
<!--
<table width="100%" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#c3d4e5">
<tr bgcolor="#FCFCFC" align="center">
<td width ="5%">No.</td>
<td width ="25%">Line������)</td>
<td width ="5%">����</td>
</tr>

  Call mainbody 

</table>
 -->
</td><td valign ="top">
<div>
<table width="100%" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#c3d4e5">
<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="25">Line</td>
<td width="10%">����</td>
<td width="5%">����</td>
<td width="5%">��λ</td>
<td width="5%">����</td>
<td width="5%">����</td>
<td width="8%">ְ��</td>
<td width="10%">������</td>
<td width="5%">���</td>
<td width="5%">����</td>
</tr>

<% 
Call maindetail
Call CloseDB
%>

</table>
</div>
</td></tr>
</table>
</body>
</html>


<%
Sub mainbody()	
   if rs.eof then
      response.write "<tr><td colspan=3>"
      response.write "��û�м�¼��"
      response.write "</td></tr>"
   else
      n=1
      while not rs.eof
      response.write "<tr bgcolor='#FCFCFC'>"
      response.write "<td align ='left'>"&n&"</td>"
      response.write "<td align ='left'><a href ='lc_count.asp?line="&rs("line")&"&model="&rs("model")&"'>"&rs("line")&"("&rs("model")&")</a></td>"
      response.write "<td align ='left'>"&rs("��λ��")&"</td>" 
      response.write "</tr>"
      rs.MoveNext
      n=n+1
      wend
   end if

End Sub


Sub maindetail()
If rs_detail.eof then
	response.write "<tr><td colspan=9>"
	response.write "û�м�¼��"
	response.write "</td></tr>"
end if

while not rs_detail.eof 
        if rs_detail("״̬")="δ��ְ" then
	response.write"<tr bgcolor='#FCFCFC'>"
        else
        response.write"<tr bgcolor='#FCACFC'>"
        end if	
	'response.write"<tr bgcolor='#FCFCFC'>"	
        response.write"<td>"
        response.write"<a href=LcRecord.asp?line="&trim(rs_detail("line"))&"&model="&trim(replace(rs_detail("����"),"","%"))&"&wogroup="&trim(rs_detail("����"))&"&workstation="&trim(rs_detail("��λ"))&">"&trim(rs_detail("Line"))&"</a></td>"
        response.write"<td>"&rs_detail("����")&"</td>"
        response.write"<td>"&rs_detail("����")&"</td>"
        response.write"<td>"&rs_detail("��λ")&"</td>"
        response.write"<td><a href=../wokermanage/emrecord.asp?CardID="&rs_detail("����")&">"&rs_detail("����")&"</a></td>"
	'response.write"<td>"&rs_detail("����")&"</td>"
        response.write"<td>"&rs_detail("����")&"</td>"
        response.write"<td>"&rs_detail("ְ��")&"</td>"
        response.write"<td>"&rs_detail("������")&"</td>"
        response.write"<td>"&rs_detail("���")&"</td>"
        response.write"<td><a href=lc_count.asp?act=cancel&liid="&rs_detail("LINEID")&">��λ</a></td>"
	response.write"</tr>"
	rs_detail.MoveNext
wend

End Sub
%>