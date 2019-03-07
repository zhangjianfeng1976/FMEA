<!--#include file="../includes/db.asp"-->
<%
sql = "SELECT * FROM t_Fmea_PJ_history where PjID ="&request("PjID")
pgs = 25

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
<title>机种管理</title>
</head>

<body bgcolor ="#FCFCFC">
<form action="His_Pj_Browse.asp">

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=11>
<b>版本修订记录列表</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="20">FMEA编号及版本</td>
<td width="10%">项目</td>
<td width="12%">名称</td>
<td width="8%">机种</td>
<td width="5%">作成者</td>
<td width="10%">过程责任</td>
<td width="10%">主要负责者</td>
<td width="8%">FMEA初始日期</td>
<td width="8%">更新日期</td>
</tr>

<% 
  call MainBody()
%>

<tr>
<td width="75%" colspan=10 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">回到首页</a>] [<a href="javascript:history.back()">返回</a>]  
</td><td  width="50%" align="right">
<% call clspage() %> 
</table>
</td>
</tr>
</table>
<% call closeDB() %>

</form>

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

	PjKey = rs("PjKey")
    Serial = rs("Serial")
    
	response.write"<tr bgcolor='#FCFCFC'>"
	response.write"<td><div>"&Pjkey&"</div>"
    response.write"<div>(<a href=His_Process_Browse.asp?cat=2&Pjkey="&Pjkey&">评价修订</a>)"
    response.write"(<a href=His_Advice_Browse.asp?cat=2&Pjkey="&Pjkey&">对应修订</a>)"
    response.write"</div></td>"
	response.write"<td>"&rs("ProjectName")&"</td>"
    response.write"<td>"&rs("ItemName")&"</td>"
    response.write"<td>"&rs("Model")&"</td>"
    response.write"<td>"&rs("Author")&"</td>"
    response.write"<td>"&rs("BurdenDept")&"</td>"
	response.write"<td>"&rs("Onduty")&"</td>"
    response.write"<td>"&rs("CreateDate")&"</td>"
    response.write"<td>"&rs("ModifyDate")&"</td>"
	response.write"</td></tr>"
	
	rs.MoveNext
	
n=n+1
wend

End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&pagecount&">尾页</a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=His_Pj_Browse.asp?page=1>首页</a>&nbsp;&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&page-1&">上页</a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=His_Pj_Browse.asp?page=1>首页</a>&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&page-1&">上页</a>&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=His_Pj_Browse.asp?page="&pagecount&">尾页</a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub

%>

