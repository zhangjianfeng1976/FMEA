<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "查询" 	                
            select case Request("cata")
              case "高级查询"
              case "模糊查询"
	    Response.Redirect "outemploybase_1.asp?Cat=0&Sc="&Request("Search")			   		
              case "工程外人员"
                    Response.Redirect "outemploybase_1.asp?Cat=1&Sc="&Request("Search")
          end select            
     Case "输入"
       Response.Redirect "add.asp"
End Select


pgs =2500
select case request("cat")
       case "" 
       sql="SELECT * FROM  v_outemploybase_1"
       pgs =100
       case 0
       sql="SELECT * FROM v_outemploybase_1 where cardid like '%"&request("Sc")& _
           "%' Or name like '%"&request("Sc")& _
           "%' Or vocation like '%"&request("Sc")& _
           "%' Or dpcode like '%"&request("Sc")& _
           "%' Or department like '%"&request("Sc")& _
           "%' Or sex like '%"&request("Sc")& _
           "%' ORDER BY cardid"    
       case 1
                  
end select

call openDB()
rs.open sql,conn,1,1
if not rs.eof then
rs.pagesize =pgs
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
<title>工位人员管理</title>
</head>

<body>
<form action="outemploybase_1.asp">


<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=10>
<b>工程外人员列表</b></td>
</tr>

<tr>
<td height="20" align='left' bgcolor="#e6f7ff" colspan="6">
快捷确认：
<a href="outemploybase_1.asp?Cat=0&Sc=A4">机器1制造部</a>&nbsp;&nbsp;
<a href="outemploybase_1.asp?Cat=0&Sc=A5">机器2制造部</a>&nbsp;&nbsp;
<a href="outemploybase_1.asp?Cat=0&Sc=A6">机器3制造部</a>&nbsp;&nbsp;
</td>
<td align='right' bgcolor="#e6f7ff" colspan ="4">
<select name='cata' class="smallinput">
<option>模糊查询</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="查询" name="OK" class="smallInput">
<%
'  response.write "<input type='submit' value='输入' name='OK' class='smallInput'>"
%>
</td>
</tr>

<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="20">工号</td>
<td width="5%">姓名</td>
<td width="8%">职称</td>
<td width="4%">性别</td>
<td width="4%">系编号</td>
<td width="30%">所属</td>
<td width="8%">入厂日期</td>
<td width="8%">离职日</td>
<td width="8%">状态</td>
<td width="8%" align="center">工程外人员考勤</td> 
</tr>

<% call MainBody() %>

<tr>
<td width="75%" colspan=10 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
<% call clspage() %>
</td><td  width="50%" align="right">
[<a href="javascript:history.back()">返回</a>]   
<% Call CloseDB() %>
</table>
</td>
</tr>
</table>
</form>
</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=10>"
	response.write "没有记录。"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
	CardID=rs("CardID")
	response.write"<tr bgcolor='#FCFCFC'>"
	response.write"<td><a href=emrecord.asp?CardID="&rs("CardID")&">" &CardID&"</a></td>"	
	response.write"<td>"&rs("name")&"</td>"
        response.write"<td>"&rs("vocation")&"</td>"
        response.write"<td>"&rs("sex")&"</td>"
        response.write"<td>"&rs("dpcode")&"</td>"
        response.write"<td>"&rs("department")&"</td>"
	response.write"<td>"&rs("entrancedate")&"</td>"
        response.write"<td>"&rs("lastwd")&"</td>"
        response.write"<td>"&rs("statuss")&"</td>"
 	response.write"<td><a href=emEdit.asp?ID="&rs("id")&">"&"报考勤</a>&nbsp;&nbsp;&nbsp;&nbsp;"
       ' response.write"<a href=add.asp?CardID="&rs("CardID")&">"&"安排入位</a>&nbsp;&nbsp;"
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&pagecount&">尾页</a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=outemploybase_1.asp?page=1>首页<a>&nbsp;&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&page-1&">上页</a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=outemploybase_1.asp?page=1>首页</a>&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&page-1&">上页</a>&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=outemploybase_1.asp?page="&pagecount&">尾页</a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub


%>
