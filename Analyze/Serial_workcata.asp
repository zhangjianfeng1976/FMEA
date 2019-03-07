<!--#include file="../includes/db.asp"-->
<%
sql="select * from v_serial_workcata P Pivot (SUM(Quant) FOR WorkCata IN ([打螺丝],[装入],[张贴],[结线],[线处理],[涂布],[清扫],[设定],[检查],[-])) AS pvt "
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
<title></title>
</head>

<body>
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=30>
<b>机种作业别分析</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="8%">机种系列</td>
<td width="4%">打螺丝</td>
<td width="4%">装入</td>
<td width="4%">张贴</td>
<td width="4%">结线</td>
<td width="4%">线处理</td>
<td width="4%">涂布</td>
<td width="4%">清扫</td>
<td width="4%">设定</td>
<td width="4%">检查</td>
<td width="4%">-</td>
<td width="7%">合计</td>
</tr>

<% call MainBody() %>


</table>
<% Call CloseDB() %>
</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=6>"
	response.write "没有记录。"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
   	response.write"<tr bgcolor='#FCFCFC'>"
    response.write"<td>"&rs("Serial")&"</td>"	
	response.write"<td>"&rs("打螺丝")&"</td>"
    i1 = rs("打螺丝")
    if isnull(i1) then i1 = 0
    a1 = a1 + i1
    response.write"<td>"&rs("装入")&"</td>"
    i2 = rs("装入")
    if isnull(i2) then i2 = 0
    a2 = a2 + i2
    response.write"<td>"&rs("张贴")&"</td>"
    i3 = rs("张贴")
    if isnull(i3) then i3 = 0
    a3 = a3 + i3
    response.write"<td>"&rs("结线")&"</td>"
    i4 = rs("结线")
    if isnull(i4) then i4 = 0
    a4 = a4 + i4
	response.write"<td>"&rs("线处理")&"</td>"
    i5 = rs("线处理")
    if isnull(i5) then i5 = 0
    a5 = a5 + i5
    response.write"<td>"&rs("涂布")&"</td>"
    i6 = rs("涂布")
    if isnull(i6) then i6 = 0
    a6 = a6 + i6
    response.write"<td>"&rs("清扫")&"</td>"
    i7 = rs("清扫")
    if isnull(i7) then i7 = 0
    a7 = a7 + i7
    response.write"<td>"&rs("设定")&"</td>"
    i8 = rs("设定")
    if isnull(i8) then i8 = 0
    a8 = a8 + i8
   	response.write"<td>"&rs("检查")&"</td>"
    i9 = rs("检查")
    if isnull(i9) then i9 = 0
    a9 = a9 + i9
    response.write"<td>"&rs("-")&"</td>"
    i10 = rs("-")
    if isnull(i10) then i10 = 0
    a10 = a10 + i10         
    response.write"<td>"&i1 + i2 + i3 + i4 +  i5 + i6 + i7 + i8 + i9 + i10&"</td>"
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend

    response.write"<tr  bgcolor='#F59595'>"
    response.write"<td>[<a href='javascript:history.back()'>返回</a>]"	
	response.write"合计</td>"      
    response.write"<td>"&a1&"</td>"
    response.write"<td>"&a2&"</td>"
    response.write"<td>"&a3&"</td>"
    response.write"<td>"&a4&"</td>"
    response.write"<td>"&a5&"</td>"
    response.write"<td>"&a6&"</td>"
    response.write"<td>"&a7&"</td>"
    response.write"<td>"&a8&"</td>"
    response.write"<td>"&a9&"</td>"
    response.write"<td>"&a10&"</td>"   
    response.write"<td>"&a1 + a2 + a3 + a4  + a5 + a6 + a7 + a8 + a9 + a10&"</td>"
	response.write"</tr>"

End Sub
%>
