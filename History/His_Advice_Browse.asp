<!--#include file="../includes/db.asp"-->
<%
pgs =2500
select case request("cat")
    case ""
     sql="SELECT * FROM t_Fmea_Advice_history Where itmID = "& request("itmID")       
    case 2
     sql="SELECT * FROM t_Fmea_Advice_history Where Pjkey = '"&request("Pjkey")&"'"
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
<title>改善措施项目一览</title>
</head>

<body bgcolor ="#FCFCFC">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=13>
<b>改善措施修订记录</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="15%" rowspan ="2">管理No.</td>
<td width="15%" rowspan ="2">建议措施</td>
<td width="5%"  rowspan ="2">责任者</td>
<td width="8%"  rowspan ="2">目标完成日</td>
<td width="18%" colspan ="5">措施结果</td>
</tr><tr bgcolor="#FCFCFC" align="center">
<td width="10%">采用措施</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RPN</td>
</tr>

<% 
call MainBody()
 %>

<tr>
<td width="75%" colspan=13 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">回到首页</a>][<a href="javascript:history.back()">返回</a>]
</td><td  width="50%" align="right">
<% call clspage() %>   
</table>
</td>
</tr>
</table>
<% call closeDB() %>
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
    AdID = rs("AdID")
	ItmID = rs("ItmID")
	response.write"<tr bgcolor='#FCFCFC'>"	
	response.write"<td>"&rs("Pjkey")&"</td>"
    response.write"<td>"&rs("AdvContent")&"</td>"	
	response.write"<td>"&rs("Rspnser")&"</td>"
    response.write"<td>"&rs("FshDate")&"</td>"
    response.write"<td>"&rs("ActionContent")&"</td>"
    response.write"<td>"&rs("ResultS")&"</td>"
    response.write"<td>"&rs("ResultO")&"</td>"
    response.write"<td>"&rs("ResultD")&"</td>"
    response.write"<td>"&rs("ResultRPN")&"</td>" 
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&pagecount&">尾页</a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=His_Advice_Browse.asp?page=1>首页</a>&nbsp;&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&page-1&">上页</a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=His_Advice_Browse.asp?page=1>首页</a>&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&page-1&">上页</a>&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=His_Advice_Browse.asp?page="&pagecount&">尾页</a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub
%>

