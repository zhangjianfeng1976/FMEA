<!--#include file="../includes/db.asp"-->
<%
select case request("cat")
  case ""
  sql="SELECT * FROM  v_dpstru order by LINE别"
  case else
  sql="SELECT * FROM  v_dpstru where LINE别 = '"&request("cat")&"'"
end select

call openDB()
rs.open sql,conn,1,1
if not rs.eof then
rs.pagesize = 1000
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
<title>人员构成</title>
</head>

<body>
<table width="80%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=7>
<b>LINE别人员构成列表</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="25%">LINE</td>
<td width="6%">部门CODE</td>
<td width="6%">在席</td>
<td width="6%">编制外</td>
<td width="6%">工程外</td>
<td width="6%">工程上</td>
<td width="6%">工程外比率</td>
</tr>

<% call MainBody() %>


</table>
<% Call CloseDB() %>
</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=5>"
	response.write "没有记录。"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
	response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<td><a href=dpstruture.asp?cat="&rs("Line别")&">"&rs("Line")&"</a></td>"	
        response.write"<td><a href=../wokermanage/browse.asp?cat=2&sc="&rs("Line别")&">"&rs("Line别")&"</a></td>"	
        i0= rs("人数")
        response.write"<td>"&i0&"</td>"
        a=a+ i0
        i1 = rs("编外")
        if isnull(i1) then 
           i1 = 0
        end if
        response.write"<td>"&i1&"</td>"
        b=b+ i1
        i2 = rs("工程外")
        if isnull(i2) then 
          i2 = 0
        end if
        response.write"<td>"&i2&"</td>"
        c=c+ i2
        response.write"<td>"&i0-i1-i2&"</td>"
        response.write"<td>"&round((i1+i2)/i0*100,2)&"%</td>"
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
        response.write"<tr  bgcolor='#F59595'>"
        response.write"<td>[<a href='javascript:history.back()'>返回</a>]</td>"
        response.write"<td>合计</td>"	   
        response.write"<td>"&a&"</td>"
        response.write"<td>"&b&"</td>"
        response.write"<td>"&c&"</td>"
        response.write"<td>"&a-b-c&"</td>"
        response.write"<td>"&round((b+c)/a*100,2)&"%</td>"
	response.write"</tr>"
End Sub
%>
