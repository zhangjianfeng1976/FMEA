<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "查询" 	                
            select case Request("cata")
              case "高级查询"
              case "模糊查询"
	           Response.Redirect "outwdairy.asp?Cat=0&Sc="&Request("Search")			   		
              case "工程外人员"
                    Response.Redirect "outwdairy.asp?Cat=1&Sc="&Request("Search")
          end select            
End Select

pgs =2500
select case request("cat")
       case "" 
       sql="SELECT * FROM  v_outwdairy order by 日期 DESC"
       pgs =25
       case 0
       sql="SELECT * FROM v_outwdairy where 工号 like '%"&request("Sc")& _
           "%' Or 姓名 like '%"&request("Sc")& _
           "%' Or 职称 like '%"&request("Sc")& _
           "%' Or 系编号 like '%"&request("Sc")& _
           "%' Or 原由 like '%"&request("Sc")& _
           "%' Or 所属 like '%"&request("Sc")& _
           "%' ORDER BY 工号"    
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
<form action="outwdairy.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr><td height="20" align='right' bgcolor="#e6f7ff">
<select name='cata' class="smallinput">
<option>模糊查询</option>

</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="查询" name="OK" class="smallInput">
</td>
</tr>
</table>
</form>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=10>
<b>工程外人员考勤状况表</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="20">日期</td>
<td width="5%">月份</td>
<td width="8%">工号</td>
<td width="4%">姓名</td>
<td width="4%">职称</td>
<td width="8%">原由</td>
<td width="25%">所属</td>
<td width="8%">系编号</td>
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

	response.write"<tr bgcolor='#FCFCFC'>"
	response.write"<td>"&rs("日期")&"</td>"
        response.write"<td>"&rs("月份")&"</td>"
        response.write"<td>"&rs("工号")&"</td>"
        response.write"<td>"&rs("姓名")&"</td>"
        response.write"<td>"&rs("职称")&"</td>"
	response.write"<td>"&rs("原由")&"</td>"
        response.write"<td>"&rs("所属")&"</td>"
        response.write"<td>"&rs("系编号")&"</td>"
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&page+1&">下页<a>&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&pagecount&">尾页<a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=outwdairy.asp?page=1>首页<a>&nbsp;&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&page-1&">上页<a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=outwdairy.asp?page=1>首页<a>&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&page-1&">上页<a>&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&page+1&">下页<a>&nbsp;"
Response.Write "<a href=outwdairy.asp?page="&pagecount&">尾页<a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub


%>
