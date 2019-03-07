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
<div align ='left'><b><font face="华文新魏" size="3" color="#CC6600">工位担当记录</font></b></div>
</td>
</tr> 

<% Call MainBody %>

<tr bgcolor="#E6F7FF">
<td colspan="12" align="left">[<a href="javascript:history.back()">返回</a>]</td>
</tr> 
</table>
</body>
</html>


<%
Sub MainBody()
sql="SELECT  * FROM  v_wmaster WHERE Line = '"&request("line")&"' AND 机种 = '"&request("model")&"' AND 工程 = '"&request("wogroup")&"' AND 工位 = '"&request("workstation")&"' order by 有效 desc"
   call openDB()
   rs.open sql,conn,1,1
	
   if rs.eof then
      response.write "<tr><td width=75% colspan=12>"
      response.write "无记录。"
      response.write "</td></tr>"
   else
      response.write "<tr  bgcolor='#e6F6F6' align ='center'>"
      response.write "<td width='8%'  height='20'>Line</td>"
      response.write "<td width='8%'>机种</td>"
      response.write "<td width='8%'>工程</td>"
      response.write "<td width='8%'>工位</td>"   
      response.write "</tr><tr>"
      response.write "<td>"&rs("line")&"</td>"
      response.write "<td>"&rs("机种")&"</td>"
      response.write "<td>"&rs("工程")&"</td>"
      response.write "<td>"&rs("工位")&"</td>"
      response.write "</tr><tr bgcolor='#e6F6F6' align ='center'>"
      response.write "<td width='8%'>有效</td>"
      response.write "<td width='8%'>版本</td>"
      response.write "<td width='8%'>工号</td>"
      response.write "<td width='8%'>姓名</td>"
      response.write "<td width='8%'>职称</td>"
      response.write "<td width='8%'>担当日</td>"
      response.write "<td width='8%'>新入职日</td>"
      response.write "<td width='8%'>离职日</td>"
      response.write "<td width='8%'>领班</td>"
      response.write "</tr>"
      while not rs.eof
      response.write "<tr>"
      response.write "<td>"&rs("有效")&"</td>"
      response.write "<td>"&rs("版本")&"</td>"
      response.write "<td>"&rs("工号")&"</td>"
      response.write "<td>"&rs("姓名")&"</td>"
      response.write "<td>"&rs("职称")&"</td>"
      response.write "<td>"&rs("担当日")&"</td>"
      response.write "<td>"&rs("新入职日")&"</td>"
      response.write "<td>"&rs("离职日")&"</td>"
      response.write "<td>"&rs("领班")&"</td>"  
      response.write "</tr>"
      rs.MoveNext
      wend
   end if
   call closeDB()
End Sub


%>