<!--#include file="../includes/db.asp"-->
<%
sql="select * from v_serial_partstype P Pivot (SUM(Quant) FOR PartsType IN ([框架件],[驱动件],[光学件],[电气件],[引导件],[导通件],[组立部件],[弹簧],[轴承],[海绵],[胶片],[卡环],[标签],[螺丝],[润滑材],[耗材],[其他],[-])) AS pvt "
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
<b>机种项目部品种类别分析</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="8%">机种系列</td>
<td width="4%">框架件</td>
<td width="4%">驱动件</td>
<td width="4%">光学件</td>
<td width="4%">电气件</td>
<td width="4%">引导件</td>
<td width="4%">导通件</td>
<td width="4%">组立部件</td>
<td width="4%">弹簧</td>
<td width="4%">轴承</td>
<td width="4%">海绵</td>
<td width="4%">胶片</td>
<td width="4%">卡环</td>
<td width="4%">标签</td>
<td width="4%">螺丝</td>
<td width="4%">润滑材</td>
<td width="4%">耗材</td>
<td width="4%">其他</td>
<td width="4%">-</td>
<td width="6%">合计</td>
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
	response.write"<td>"&rs("框架件")&"</td>"
    i1 = rs("框架件")
    if isnull(i1) then i1 = 0
    a1 = a1 + i1
    response.write"<td>"&rs("驱动件")&"</td>"
    i2 = rs("驱动件")
    if isnull(i2) then i2 = 0
    a2 = a2 + i2
    response.write"<td>"&rs("光学件")&"</td>"
    i3 = rs("光学件")
    if isnull(i3) then i3 = 0
    a3 = a3 + i3
    response.write"<td>"&rs("电气件")&"</td>"
    i4 = rs("电气件")
    if isnull(i4) then i4 = 0
    a4 = a4 + i4
	response.write"<td>"&rs("引导件")&"</td>"
    i5 = rs("引导件")
    if isnull(i5) then i5 = 0
    a5 = a5 + i5
    response.write"<td>"&rs("导通件")&"</td>"
    i6 = rs("导通件")
    if isnull(i6) then i6 = 0
    a6 = a6 + i6
    response.write"<td>"&rs("组立部件")&"</td>"
    i7 = rs("组立部件")
    if isnull(i7) then i7 = 0
    a7 = a7 + i7
    response.write"<td>"&rs("弹簧")&"</td>"
    i8 = rs("弹簧")
    if isnull(i8) then i8 = 0
    a8 = a8 + i8
   	response.write"<td>"&rs("轴承")&"</td>"
    i9 = rs("轴承")
    if isnull(i9) then i9 = 0
    a9 = a9 + i9
    response.write"<td>"&rs("海绵")&"</td>"
    i10 = rs("海绵")
    if isnull(i10) then i10 = 0
    a10 = a10 + i10
    response.write"<td>"&rs("胶片")&"</td>"
    i11 = rs("胶片")
    if isnull(i11) then i11 = 0
    a11 = a11 + i11
    response.write"<td>"&rs("卡环")&"</td>"
    i12 = rs("卡环")
    if isnull(i12) then i12 = 0
    a12 = a12 + i12
	response.write"<td>"&rs("标签")&"</td>"
    i13 = rs("标签")
    if isnull(i13) then i13 = 0
    a13 =a13 + i13
    response.write"<td>"&rs("螺丝")&"</td>"
    i14 = rs("螺丝")
    if isnull(i14) then i14 = 0
    a14 = a14 + i14
    response.write"<td>"&rs("润滑材")&"</td>"
    i15 = rs("润滑材")
    if isnull(i15) then i15 = 0
    a15 = a15 + i15
    response.write"<td>"&rs("耗材")&"</td>"
    i16 = rs("耗材")
    if isnull(i16) then i16 = 0
    a16 = a16 + i16
    response.write"<td>"&rs("其他")&"</td>"
    i17 = rs("其他")
    if isnull(i17) then i17 = 0
    a17 = a17 + i17
    response.write"<td>"&rs("-")&"</td>"
    i18 = rs("-")
    if isnull(i18) then i18 = 0
    a18 = a18 + i18

    response.write"<td>"&i1 + i2 + i3 + i4  + i5 + i6 + i7 + i8 + i9 + i10 + i11 + i12 + i13 + i14  + i15 + i16 + i17 + i18 &"</td>"
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
    response.write"<td>"&a11&"</td>"
    response.write"<td>"&a12&"</td>"
    response.write"<td>"&a13&"</td>"
    response.write"<td>"&a14&"</td>"
    response.write"<td>"&a15&"</td>"
    response.write"<td>"&a16&"</td>"
    response.write"<td>"&a17&"</td>"
    response.write"<td>"&a18&"</td>"
    response.write"<td>"&a1 + a2 + a3 + a4  + a5 + a6 + a7 + a8 + a9 + a10 + a11 + a12 + a13 + a14  + a15 + a16 + a17 + a18&"</td>"
	response.write"</tr>"

End Sub
%>
