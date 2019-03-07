<!--#include file="../includes/db.asp"-->
<%
pgs =2500
select case request("cat")
    case ""
        sql="SELECT * FROM t_Fmea_Item_history where  ItmID = "&request("ItmID")
    case 2
        sql="SELECT * FROM t_Fmea_Item_history where  Pjkey = '"&request("Pjkey")&"' order by subpjname,pageno,workqueue"
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
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE2 {color: #0000FF}
-->
</style>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>机种管理项目修订一览</title>
</head>

<body bgcolor ="#FCFCFC">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=18>
<b><% = request("Pjkey") %>作业内容评价修订一览</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="2%">编号</td>
<td width="2%">顺序</td>
<td width="5%" height="20">工程名称</td>
<td width="5%">标准书页码</td>
<td width="7%">部品编号</td>
<td width="10%">部品名称及种类</td>
<td width="8%">作业分类与内容</td>
<td width="10%">潜在缺陷模式</td>
<td width="10%">潜在缺陷后果</td>
<td width="7%">现行工程设计及过程控制</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RPN</td>
<td width="2%">措施</td> 
</tr>

<% 
call MainBody() 
%>

<tr>
<td width="75%" colspan=18 height="30" bgcolor="#EFEFEF" align='center'>
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
	ItmID = rs("ItmID")
	response.write"<tr bgcolor='#FCFCFC'>"
	response.write"<td>"&rs("Pjkey")&"</td>"
    response.write"<td>"&rs("WorkQueue")&"</td>"
    response.write"<td>"&rs("SubPjName")&"</td>"
    response.write"<td>"&rs("PageNo")&"</td>"
	response.write"<td>"&rs("PartsNo")&"</a></td>"
    response.write"<td><div>("&rs("PartsType")&")</div>"&rs("PartsName")&"</td>"
    response.write"<td><div>("&rs("WorkCata")&")</div><div>"&rs("WorkContent")&"</div></td>"
    response.write"<td>"&rs("NGMode")&"</td>"
	response.write"<td>"&rs("NGEffect")&"</td>"
    response.write"<td>"&rs("ProDesign")&"</td>"
    response.write"<td align ='center'>"&rs("Sev")&"</td>"
    response.write"<td align ='center'>"&rs("Occ")&"</td>"
    response.write"<td align ='center'>"&rs("Det")&"</td>"
    
    vRPN = trim(rs("RPN"))
    if vRPN > 2 Then
        response.write"<td bgcolor = '#DDFFAA'><span class='STYLE1'>&nbsp;"& vRPN &"</span></td>"    
        response.write"<td bgcolor = '#DDFFAA'><a href=his_Advice_Browse.asp?ItmID="&ItmID&">对应</a></td>" 
        response.write"</tr>"
    else
        response.write"<td>"& vRPN &"</td>" 
        response.write"<td></td>"
        response.write"</tr>"
    end if
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&pagecount&">尾页</a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=His_Process_Browse.asp?page=1>首页</a>&nbsp;&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&page-1&">上页</a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=His_Process_Browse.asp?page=1>首页</a>&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&page-1&">上页</a>&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=His_Process_Browse.asp?page="&pagecount&">尾页</a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub
%>

