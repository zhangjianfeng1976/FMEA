<!--#include file="../includes/db.asp"-->
<%
 sql="SELECT * FROM  v_drrate order by 科别"
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
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=14>
<b>直间比人员构成列表</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="10%">科别</td>
<td width="4%" bgcolor=#e6F6F6>人员总数</td>
<td width="4%">领班1级</td>
<td width="4%">领班2级</td>
<td width="4%">组长1级</td>
<td width="4%">组长2级</td>
<td width="4%" bgcolor=#e6F6F6>间接人员</td>
<td width="4%">MEISTER B</td>
<td width="4%">MEISTER C</td>
<td width="4%">员工2级</td>
<td width="4%">员工3级</td>
<td width="4%" bgcolor=#e6F6F6>直接人员</td>
<td width="4%">组长员工比</td>
<td width="4%">直间比</td>
</tr>

<% call MainBody() %>


</table>
<% Call CloseDB() %>
</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=14>"
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
        response.write"<td>"&depart&"</td>"
        response.write"<td bgcolor=#e6F6F6>"&rs("人员数")&"</td>"
        a = a +	rs("人员数")
	response.write"<td>"&rs("领班1级")&"</td>"
        b= b + rs("领班1级")
        response.write"<td>"&rs("领班2级")&"</td>"
        c= c + rs("领班2级")
        response.write"<td>"&rs("组长1级")&"</td>"
        d= d + rs("组长1级")
        response.write"<td>"&rs("组长2级")&"</td>"
        e= e + rs("组长2级")
        response.write"<td bgcolor=#e6F6F6>"&rs("间接人员")&"</td>"
        f= f + rs("间接人员")
        response.write"<td>"&rs("MEISTER_B")&"</td>"
        g= g + rs("MEISTER_B")
        response.write"<td>"&rs("MEISTER_C")&"</td>"
        h= h + rs("MEISTER_C")
        response.write"<td>"&rs("员工2级")&"</td>"
        i= i + rs("员工2级")
        response.write"<td>"&rs("员工3级")&"</td>"
        j= j + rs("员工3级")
        response.write"<td bgcolor=#e6F6F6>"&rs("直接人员")&"</td>"
        k= k + rs("直接人员")

        x = rs("组长员工比")

     '   if x <> 0 then
     '    y = 100/x
     '   else 
     '    y = x
     '   end if

        response.write"<td>"&x&"%</td>"
        xdirect = rs("间接人员")
        if xdirect = 0 then
        response.write "<td>None</td>"
        else
        response.write"<td>"&round(rs("直接人员")/xdirect,2)&"</td>"
	end if
        response.write"</tr>"
	rs.MoveNext
n=n+1
wend
       response.write"<tr bgcolor='#F59595'>"
       response.write"<td>合计</td>"
       response.write"<td>"&a&"</td>"	
       response.write"<td>"&b&"</td>"
       response.write"<td>"&c&"</td>"
       response.write"<td>"&d&"</td>"
       response.write"<td>"&e&"</td>"
       response.write"<td>"&f&"</td>"
       response.write"<td>"&g&"</td>"
       response.write"<td>"&h&"</td>"
       response.write"<td>"&i&"</td>"
       response.write"<td>"&j&"</td>"
       response.write"<td>"&k&"</td>"
       response.write"<td>"&round((g+ h + i + j)/(d + e),2)&"<div>(员工组长比)</div></td>"
       response.write"<td>"&round(k/f,2)&"</td>"
       response.write"</tr>"
End Sub
%>
