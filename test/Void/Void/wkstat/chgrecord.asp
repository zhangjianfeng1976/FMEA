<!--#include file="includes/db.asp"-->
<%
Select case request("Action")
       case "Cancel"
            if trim(request("truename"))<>trim(Request.ServerVariables("REMOTE_HOST")) then 
              response.write "�ü�¼����������ģ���û��Ȩ��ȡ��[<a href='javascript:history.back()'>����</a>]"
              response.end
            end if
              sql = "Update wt_wdairy set drstatus = 'Cancel' where ID = "&request("ID")
              call openDB()
              conn.execute(sql)
              call closeDB()
end select
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Includes/Style.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<body>
<form action="chgrecord.asp">
<table width="100%" border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr bgcolor="#E6F7FF">
<td width="50%" bgcolor="#E6F7FF" colspan="13">
<div align ='left'><b><font face="������κ" size="3" color="#CC6600">������λ��¼</font></b></div>
</td>
</tr> 

<tr bgcolor="#FCFCFC" align="center">
<td width="2%" height="20">No.</td>
<td width="6%">����</td>
<td width="4%">Line</td>
<td width="8%">����</td>
<td width="6%">����</td>
<td width="6%">��λ</td>
<td width="6%">ԭ��ҵ�߹���</td>
<td width="6%">����</td>
<td width="6%">ԭ��</td>
<td width="6%">��λ�߹���</td>
<td width="6%">��λ</td>
<td width="6%">Ŀ��</td>
<td width="4%">����</td>
</tr>

<% Call dairy %>

</table>
</form>
</body>
</html>


<%
Sub dairy()
sql="SELECT  * FROM v_wdairy where ���� = '"&date()&"' order by line"
   call openDB()
   rs.open sql,conn,1,1
	
   if rs.eof then
      response.write "<tr><td width=75% colspan=12>"
      response.write "��û�м�¼��"
      response.write "</td></tr>"
   else
      n =1
      while not rs.eof
      line = rs("line")
      response.write "<tr align ='center'>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&n&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("����")&"</td>"
      response.write "<td width='9%' height='16' bgcolor='#FCFCFC'>"      
      response.write line
      response.write "</td>"
      response.write "<td bgcolor='#FCFCFC'>"&rs("����")&"</td>"
      response.write "<td bgcolor='#FCFCFC'>"&rs("����")&"</td>"
      response.write "<td bgcolor='#FCFCFC'>"&rs("��λ")&"</td>" 
      response.write "<td bgcolor='#FCFCFC'>"&rs("ԭ��ҵ�߹���")&"</td>"
      response.write "<td bgcolor='#FCFCFC'>"&rs("����")&"</td>"
      response.write "<td bgcolor='#FCFCFC'>"&rs("ԭ��")&"</td>"
      response.write "<td bgcolor='#FCFCFC'>"&rs("��λ�߹���")&"</td>"
      response.write "<td bgcolor='#e6F6F6'>"&rs("��λ")&"</td>"
      response.write "<td bgcolor='#FCFCFC'>"&rs("Ŀ��")&"</td>"
      response.write "<td bgcolor='#FCFCFC'><a href=chgrecord.asp?Action=Cancel&truename="&rs("truename")&"&ID="&rs("id")&">ȡ��</a></td>"
      response.write "</tr>"

      rs.MoveNext
      n= n+1
      wend
   end if
   call closeDB()
End Sub


%>