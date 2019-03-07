<!--#include file="../includes/db.asp"-->
<%
sql="SELECT  * FROM v_lc_count"
Call OpenDB
rs.open sql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Includes/Style.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<body>
<table width="100%" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#c3d4e5">
<tr bgcolor="#FCFCFC" align="center">
<td width ="5%">No.</td>
<td width ="25%">Line（机种)</td>
<td width ="5%">数量</td>
</tr>

<%  
Call mainbody 
Call CloseDB
%>

</table>
</td></tr>
</table>
</body>
</html>

<%
Sub mainbody()	
   if rs.eof then
      response.write "<tr><td colspan=3>"
      response.write "还没有记录。"
      response.write "</td></tr>"
   else
      n=1
      while not rs.eof
      response.write "<tr bgcolor='#FCFCFC'>"
      response.write "<td align ='left'>"&n&"</td>"
      response.write "<td align ='left'><a href ='lc_count.asp?line="&trim(rs("line"))&"&model="&trim(rs("model"))&"' target = _parent>"&trim(rs("line"))&"("&trim(rs("model"))&")</a></td>"
      response.write "<td align ='left'>"&rs("工位数")&"</td>" 
      response.write "</tr>"
      rs.MoveNext
      n=n+1
      wend
   end if

End Sub
%>