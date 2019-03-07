<!--#include file="../includes/db.asp"-->
<%
'sql="select * from v_RI_Cnt P Pivot (sum(cnt) FOR RPN IN 
'([4.0],[3.6],[3.3],[3.2],[3.0],[2.9],[2.6],[2.5],[2.3],[2.1],[2.0],[1.8],[1.6],[1.4],[1.3],[1.0])) AS pvt "

sql ="select * from v_RI_Cnt_CaseWhen order by Serial"
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
<b>机种项目评分表</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="2%">No.</td>
<td width="8%">机种系列</td>
<td width="2%" bgcolor ="#FFEFEF">4.0</td>
<td width="2%" bgcolor ="#FFEFEF">3.6</td>
<td width="2%" bgcolor ="#FFEFEF">3.3</td>
<td width="2%" bgcolor ="#FFEFEF">3.2</td>
<td width="2%" bgcolor ="#FFFF9F">3.0</td>
<td width="2%" bgcolor ="#FFFF9F">2.9</td>
<td width="2%" bgcolor ="#FFFF9F">2.6</td>
<td width="2%" bgcolor ="#FFFF9F">2.5</td>
<td width="2%" bgcolor ="#FFFF9F">2.3</td>
<td width="2%" bgcolor ="#FFFF9F">2.1</td>
<td width="2%">2.0</td>
<td width="2%">1.8</td>
<td width="2%">1.6</td>
<td width="2%">1.4</td>
<td width="2%">1.3</td>
<td width="2%">1.0</td>
<td width="4%">合计</td>
<td width="4%">2.1以上</td>
<td width="4%">比率</td>
</tr>

<% 
call MainBody()

%>


</table>
<% Call CloseDB() %>
</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=20>"
	response.write "没有记录。"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
        response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<td>"&n&"</td>"
        response.write"<td>"&rs("Serial")&"</td>"	
       response.write"<td bgcolor='#FFEFEF'>"&rs("4.0")&"</td>"
        i1 = rs("4.0")
        if isnull(i1) then  i1 = 0
        a1=a1+ i1
        response.write"<td bgcolor='#FFEFEF'>"&rs("3.6")&"</td>"
        i2 = rs("3.6")
        if isnull(i2) then i2 = 0
        a2=a2+ i2
        response.write"<td bgcolor='#FFEFEF'>"&rs("3.3")&"</td>"
        i3 = rs("3.3")
        if isnull(i3) then i3 = 0
        a3=a3+ i3
        response.write"<td bgcolor='#FFEFEF'>"&rs("3.2")&"</td>"
        i4 = rs("3.2")
        if isnull(i4) then i4 = 0
        a4=a4+ i4
        response.write"<td bgcolor='#FFFF9F'>"&rs("3.0")&"</td>"
        i5 = rs("3.0")
        if isnull(i5) then  i5 = 0
        a5=a5 + i5
        response.write"<td bgcolor ='#FFFF9F'>"&rs("2.9")&"</td>"
        i6 = rs("2.9")
        if isnull(i6) then  i6 = 0
        a6=a6+ i6
        response.write"<td bgcolor ='#FFFF9F'>"&rs("2.6")&"</td>"
        i7 = rs("2.6")
        if isnull(i7) then  i7 = 0
        a7=a7+ i7
        response.write"<td bgcolor ='#FFFF9F'>"&rs("2.5")&"</td>"
        i8 = rs("2.5")
        if isnull(i8) then i8 = 0
        a8=a8 + i8
        response.write"<td bgcolor ='#FFFF9F'>"&rs("2.3")&"</td>"
        i9 = rs("2.3")
        if isnull(i9) then  i9 = 0
        a9=a9+ i9
        response.write"<td bgcolor ='#FFFF9F'>"&rs("2.1")&"</td>"
        i10 = rs("2.1")
        if isnull(i10) then  i10 = 0
        a10 = a10 + i10
        response.write"<td>"&rs("2.0")&"</td>"
        i11 = rs("2.0")
        if isnull(i11) then  i11 = 0
        a11 = a11 + i11
        response.write"<td>"&rs("1.8")&"</td>"
        i12 = rs("1.8")
        if isnull(i12) then  i12 = 0
         a12 = a12 + i12
        response.write"<td>"&rs("1.6")&"</td>"
        i13 = rs("1.6")
        if isnull(i13) then i13 = 0
        a13 =a13 + i13
        response.write"<td>"&rs("1.4")&"</td>"
        i14 = rs("1.4")
        if isnull(i14) then  i14 = 0
        a14 = a14 + i14
        response.write"<td>"&rs("1.3")&"</td>"
        i15 = rs("1.3")
        if isnull(i15) then i15 = 0
        a15 = a15 + i15
        response.write"<td>"&rs("1.0")&"</td>"
        i16 = rs("1.0")
        if isnull(i16) then i16 = 0
        a16 = a16 + i16
        response.write"<td bgcolor ='#eFfFeF'>"&i1 + i2 + i3 + i4  + i5 + i6 + i7 + i8 + i9 + i10 + i11 + i12 + i13 + i14  + i15 + i16  &"</td>"
        response.write"<td>"&i1 + i2 + i3 + i4  + i5 + i6 + i7 + i8 + i9 + i10  &"</td>"
        response.write"<td>"&round((i1 + i2 + i3 + i4  + i5 + i6 + i7 + i8 + i9 + i10)/ (i1 + i2 + i3 + i4  + i5 + i6 + i7 + i8 + i9 + i10 + i11 + i12 + i13 + i14  + i15 + i16),3)*100 &"%</td>"
        response.write"</tr>"
        rs.MoveNext
n=n+1
wend

        response.write"<tr  bgcolor='#F59595'>"
        response.write"<td colspan =2>[<a href='javascript:history.back()'>返回</a>]"	
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

        response.write"<td>"&a1 + a2  + a3 + a4  + a5 + a6 + a7 + a8 + a9 + a10 + a11 + a12 + a13 + a14  + a15 + a16  &"</td>"
        response.write"<td>"&a1 + a2  + a3 + a4  + a5 + a6 + a7 + a8 + a9 + a10  &"</td>"
        response.write"<td>"&round((a1 + a2  + a3 + a4  + a5 + a6 + a7 + a8 + a9 + a10)/(a1 + a2  + a3 + a4  + a5 + a6 + a7 + a8 + a9 + a10 + a11 + a12 + a13 + a14  + a15 + a16),3)*100 &"%</td>"
        response.write"</tr>"

End Sub
%>
