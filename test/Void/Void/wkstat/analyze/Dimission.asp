<!--#include file="../includes/db.asp"-->
<%
sql="select * from sub_v_dimission P Pivot (SUM(TTL) FOR DEP IN ([A41],[A42],[A48],[A54],[A55],[A61],[A62],[A65])) AS pvt "
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
<td height="30" align='center' bgcolor="#E6F7FF" colspan=10>
<b>离职人员推移</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="5%">月份</td>
<td width="10%">机器1制造1科</td>
<td width="10%">机器1制造2科</td>
<td width="10%">MFP1制造推进科</td>
<td width="10%">机器2制造1科</td>
<td width="10%">机器2制造2科</td>
<td width="10%">机器3制造1科</td>
<td width="10%">机器3制造2科</td>
<td width="10%">PRT制造推进科</td>
<td width="10%">合计</td>
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
        response.write"<td>"&rs("MONTHD")&"</td>"	
	response.write"<td>"&rs("A41")&"</td>"
        i1 = rs("A41")
        if isnull(i1) then 
           i1 = 0
         end if
         a1=a1+ i1
        response.write"<td>"&rs("A42")&"</td>"
        i2 = rs("A42")
        if isnull(i2) then 
           i2 = 0
         end if
         a2=a2+ i2    
        response.write"<td>"&rs("A48")&"</td>"
        i3 = rs("A48")
        if isnull(i3) then 
           i3 = 0
         end if
         a3=a3+ i3 
        response.write"<td>"&rs("A54")&"</td>"   
        i4 = rs("A54")
        if isnull(i4) then 
           i4 = 0
         end if
         a4=a4+ i4  
        response.write"<td>"&rs("A55")&"</td>"
        i5 = rs("A55")
        if isnull(i5) then 
           i5 = 0
         end if
         a5=a5+ i5 
        response.write"<td>"&rs("A61")&"</td>"
        i6 = rs("A61")
        if isnull(i6) then 
           i6 = 0
         end if
         a6=a6 + i6 
        response.write"<td>"&rs("A62")&"</td>"
        i7 = rs("A62")
        if isnull(i7) then 
           i7 = 0
         end if
         a7 =a7 + i7 
        response.write"<td>"&rs("A65")&"</td>"
        i8 = rs("A65")
        if isnull(i8) then 
           i8 = 0
         end if
         a8 =a8 + i8
        response.write"<td>"&i1+i2+i3+i4+i5+i6+i7+i8&"</td>"
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
        response.write"<td>"&a1+a2+a3+a4+a5+a6+a7+a8&"</td>"
	response.write"</tr>"

End Sub
%>
