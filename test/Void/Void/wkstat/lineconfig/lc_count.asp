<!--各LINE工位数-->
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
sql1="SELECT * FROM  v_wmaster where 有效 ='Y' and line = '"&request("line")&"' and 机种 = '"&request("model")&"'"
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
<b>各LINE/机种工程配置表</b> 
</td></tr> 
<tr><td width="300" valign="top">
<iframe src='lc_count_in.asp' name='container1' width='100%' marginwidth='1' height='470' marginheight='1' frameborder='0' scrolling='auto'></iframe>
<!--
<table width="100%" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#c3d4e5">
<tr bgcolor="#FCFCFC" align="center">
<td width ="5%">No.</td>
<td width ="25%">Line（机种)</td>
<td width ="5%">数量</td>
</tr>

  Call mainbody 

</table>
 -->
</td><td valign ="top">
<div>
<table width="100%" border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#c3d4e5">
<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="25">Line</td>
<td width="10%">机种</td>
<td width="5%">工程</td>
<td width="5%">工位</td>
<td width="5%">工号</td>
<td width="5%">姓名</td>
<td width="8%">职称</td>
<td width="10%">担当日</td>
<td width="5%">领班</td>
<td width="5%">操作</td>
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
      response.write "还没有记录。"
      response.write "</td></tr>"
   else
      n=1
      while not rs.eof
      response.write "<tr bgcolor='#FCFCFC'>"
      response.write "<td align ='left'>"&n&"</td>"
      response.write "<td align ='left'><a href ='lc_count.asp?line="&rs("line")&"&model="&rs("model")&"'>"&rs("line")&"("&rs("model")&")</a></td>"
      response.write "<td align ='left'>"&rs("工位数")&"</td>" 
      response.write "</tr>"
      rs.MoveNext
      n=n+1
      wend
   end if

End Sub


Sub maindetail()
If rs_detail.eof then
	response.write "<tr><td colspan=9>"
	response.write "没有记录。"
	response.write "</td></tr>"
end if

while not rs_detail.eof 
        if rs_detail("状态")="未离职" then
	response.write"<tr bgcolor='#FCFCFC'>"
        else
        response.write"<tr bgcolor='#FCACFC'>"
        end if	
	'response.write"<tr bgcolor='#FCFCFC'>"	
        response.write"<td>"
        response.write"<a href=LcRecord.asp?line="&trim(rs_detail("line"))&"&model="&trim(replace(rs_detail("机种"),"","%"))&"&wogroup="&trim(rs_detail("工程"))&"&workstation="&trim(rs_detail("工位"))&">"&trim(rs_detail("Line"))&"</a></td>"
        response.write"<td>"&rs_detail("机种")&"</td>"
        response.write"<td>"&rs_detail("工程")&"</td>"
        response.write"<td>"&rs_detail("工位")&"</td>"
        response.write"<td><a href=../wokermanage/emrecord.asp?CardID="&rs_detail("工号")&">"&rs_detail("工号")&"</a></td>"
	'response.write"<td>"&rs_detail("工号")&"</td>"
        response.write"<td>"&rs_detail("姓名")&"</td>"
        response.write"<td>"&rs_detail("职称")&"</td>"
        response.write"<td>"&rs_detail("担当日")&"</td>"
        response.write"<td>"&rs_detail("领班")&"</td>"
        response.write"<td><a href=lc_count.asp?act=cancel&liid="&rs_detail("LINEID")&">缩位</a></td>"
	response.write"</tr>"
	rs_detail.MoveNext
wend

End Sub
%>