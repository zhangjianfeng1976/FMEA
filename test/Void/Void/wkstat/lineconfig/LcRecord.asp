<!--#include file="../includes/db.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Includes/Style.css" rel="stylesheet" type="text/css">
<title></title>
<base target="MainWindow">
</head>

<body>
<table width="100%" border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr bgcolor="#E6F7FF">
<td width="100%" bgcolor="#E6F7FF" colspan="12">
<div align ='left'><b><font face="������κ" size="3" color="#CC6600">��λ������¼</font></b></div>
</td>
</tr> 

<% Call MainBody %>

<tr bgcolor="#E6F7FF">
<td colspan="12" align="left">[<a href="javascript:history.back()">����</a>]</td>
</tr> 
</table>
</body>
</html>


<%
Sub MainBody()
sql="SELECT  * FROM  v_wmaster WHERE Line = '"&request("line")&"' AND ���� = '"&request("model")&"' AND ���� = '"&request("wogroup")&"' AND ��λ = '"&request("workstation")&"' order by ��Ч desc"
   call openDB()
   rs.open sql,conn,1,1
	
   if rs.eof then
      response.write "<tr><td width=75% colspan=12>"
      response.write "�޼�¼��"
      response.write "</td></tr>"
   else
      response.write "<tr  bgcolor='#e6F6F6' align ='center'>"
      response.write "<td width='8%'  height='20'>Line</td>"
      response.write "<td width='8%'>����</td>"
      response.write "<td width='8%'>����</td>"
      response.write "<td width='8%'>��λ</td>"   
      response.write "</tr><tr>"
      response.write "<td>"&rs("line")&"</td>"
      response.write "<td>"&rs("����")&"</td>"
      response.write "<td>"&rs("����")&"</td>"
      response.write "<td>"&rs("��λ")&"</td>"
      response.write "</tr><tr bgcolor='#e6F6F6' align ='center'>"
      response.write "<td width='8%'>��Ч</td>"
      response.write "<td width='8%'>�汾</td>"
      response.write "<td width='8%'>����</td>"
      response.write "<td width='8%'>����</td>"
      response.write "<td width='8%'>ְ��</td>"
      response.write "<td width='8%'>������</td>"
      response.write "<td width='8%'>����ְ��</td>"
      response.write "<td width='8%'>��ְ��</td>"
      response.write "<td width='8%'>���</td>"
      response.write "</tr>"
      while not rs.eof
      response.write "<tr>"
      response.write "<td>"&rs("��Ч")&"</td>"
      response.write "<td>"&rs("�汾")&"</td>"
      response.write "<td>"&rs("����")&"</td>"
      response.write "<td>"&rs("����")&"</td>"
      response.write "<td>"&rs("ְ��")&"</td>"
      response.write "<td>"&rs("������")&"</td>"
      response.write "<td>"&rs("����ְ��")&"</td>"
      response.write "<td>"&rs("��ְ��")&"</td>"
      response.write "<td>"&rs("���")&"</td>"  
      response.write "</tr>"
      rs.MoveNext
      wend
   end if
   call closeDB()
End Sub


%>