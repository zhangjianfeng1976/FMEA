<!--#include file="../includes/db.asp"-->
<%
select case request("cat")
  case ""
  sql="SELECT * FROM  v_linestruture order by 科别,职称"
  case 1
  sql="SELECT * FROM  v_linestruture where 科别 = '"&request("ka")&"' order by 科别,职称"
  case 2
  sql="SELECT * FROM  v_linestruture where 职称 like '%"&request("ka")&"%' order by 科别,职称"
end select

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
<title>人员构成</title>
</head>

<body>
<table width="70%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=9>
<b>人员构成列表</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="10%">科别</td>
<td width="10%">职位分布</td>
<td width="4%">在席</td>
<td width="4%">编制外</td>
<td width="4%">工程外</td>
<td width="4%">工程上</td>
<td width="4%">工程外比率</td>
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
      select case rs("科别")
            case "A41"
              depart ="机器1制造部制造1科"
            case "A42"
              depart ="机器1制造部制造2科"
            case "A48"
              depart ="MFP1制造部制造推进科"
            case "A54"
              depart ="机器2制造部制造1科"
            case "A55"
              depart ="机器2制造部制造2科"
            case "A61"
              depart ="机器3制造部制造1科"
            case "A62"
              depart ="机器3制造部制造2科"
            case "A91"
              depart = "制造推进科"
            case else
              depart ="未整理"
      end select

	response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<td><a href=wkbrow.asp?cat=1&ka="&rs("科别")&">"&depart&"</a></td>"	
	response.write"<td><a href=wkbrow.asp?cat=2&ka="&rs("职称")&">"&rs("职称")&"</a></td>"
        response.write"<td>"&rs("在席")&"</td>"
        i0 = rs("在席")
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
        response.write"<td>"&i0 - i1  -i2 &"</td>"
        response.write"<td>"&round((i1+i2)/i0*100,2)&"%</td>"
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend

        response.write"<tr  bgcolor='#F59595'>"
        response.write"<td>[<a href='javascript:history.back()'>返回</a>] </td>"	
	response.write"<td>合计</td>"      
        response.write"<td>"&a&"</td>"
        response.write"<td>"&b&"</td>"
        response.write"<td>"&c&"</td>"
        response.write"<td>"&a-b-c&"</td>"
        response.write"<td>"&round((b+c)/a*100,2)&"%</td>"
	response.write"</tr>"
End Sub
%>
