<!--#include file="../includes/db.asp"-->
<%
LiID=request("liid")
if request("act") = "cancel" then
'response.write "<script language='JavaScript'>alert('缩位!');</script>" 
sqlx = "Update wt_lineconfig set wkstatus = 'N' where LineId = "&LiID
call openDB()
conn.execute(sqlx)
call closeDB()
end if

Select Case Request("OK") 
       Case "查询" 	                
            select case Request("cata")
              case "高级查询"
              case "模糊查询"                  
	            Response.Redirect "Lcbrowse_ld.asp?Cat=0&Sc="&Request("Search")       		   	
              case "按Line查询"                  
                    Response.Redirect "Lcbrowse_ld.asp?Cat=1&Sc="&Request("Search")
              case "按机种查询"                 
                    Response.Redirect "Lcbrowse_ld.asp?Cat=2&Sc="&Request("Search")
              case "按工程查询"                
                    Response.Redirect "Lcbrowse_ld.asp?Cat=3&Sc="&Request("Search")
              case "按工位查询"                 
                    Response.Redirect "Lcbrowse_ld.asp?Cat=4&Sc="&Request("Search")              		
          end select
       Case "所有工位查询" 
             Response.Redirect "Lcbrowse.asp?Sc="&Request("Search")           
End Select

pgs =500
select case request("cat")
       case "" 
       sql="SELECT * FROM  v_wmaster where 有效 ='Y' and 领班 = '"&request.Cookies("leader")&"' order by LINE,机种,工程,工位"
       pgs =200
       case 0
       sql="SELECT * FROM v_wmaster where 有效 ='Y' and 领班 = '"&request.Cookies("leader")&"' and (line like '%"&request("Sc")& _
           "%' Or 机种 like '%"&request("Sc")& _
           "%' Or 工程 like '%"&request("Sc")& _
           "%' Or 工位 like '%"&request("Sc")& _
           "%' Or 工号 like '%"&request("Sc")& _
           "%' Or 姓名 like '%"&request("Sc")& _
           "%') order by LINE,机种,工程,工位"     
       case 1
         sql="SELECT * FROM v_wmaster where 有效 ='Y' and 领班 = '"&request.Cookies("leader")&"' and line = '"&request("Sc")&"' order by 机种"
       case 2
         sql="SELECT * FROM v_wmaster where 有效 ='Y' and 领班 = '"&request.Cookies("leader")&"' and 机种 = '"&request("Sc")&"' order by line"
       case 3
         sql="SELECT * FROM v_wmaster where 有效 ='Y' and 领班 = '"&request.Cookies("leader")&"' and 工程 = '"&request("Sc")&"' order by line"  
       case 4
         sql="SELECT * FROM v_wmaster where 有效 ='Y' and 领班 = '"&request.Cookies("leader")&"' and 工位 = '"&request("Sc")&"' order by line" 
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
<form action="Lcbrowse_ld.asp">
<table  width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=11>
<table width="100%"><tr><td>
<b>[<%=request.Cookies("leader") %>]管理范围工程安排一览</b>
[<a href=lcbrowse_ld.asp>显示全部工位</a>]</td>
<td align='right' bgcolor="#e6f7ff" colspan ='10'>
<select name='cata' class="smallinput">
<option>模糊查询</option>
<option>按Line查询</option>
<option>按机种查询</option>
<option>按工程查询</option>
<option>按工位查询</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="查询" name="OK" class="smallInput">
<input type="submit" value="所有工位查询" name="OK" class="smallInput"> 
【<a href="../includes/LineConfigBatch.xls">工位人员输入</a>】
</td></tr>

</table>
</td></tr>

<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="25">Line</td>
<td width="15%">机种</td>
<td width="5%">工程</td>
<td width="8%">工位</td>
<td width="5%">工号</td>
<td width="5%">姓名</td>
<td width="5%">职称</td>
<td width="8%">担当日</td>
<td width="20%" align="center">变更安排</td>
</tr>

<% 

call MainBody()

%>

<tr>
<td width="100%" colspan=11 height="30" bgcolor="#EFEFEF" align='center'>
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
	response.write "<tr><td colspan=6>"
	response.write "没有记录。"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
	response.write"<tr bgcolor='#FCFCFC'>"	
        response.write"<td><a href=lcbrowse_ld.asp?cat=1&sc="&rs("Line")&">"&rs("Line")&"</a></td>"
        response.write"<td><a href=lcbrowse_ld.asp?cat=2&sc="&rs("机种")&">"&rs("机种")&"</a></td>"
        response.write"<td><a href=lcbrowse_ld.asp?cat=3&sc="&rs("工程")&">"&rs("工程")&"</a></td>"
        response.write"<td><a href=lcbrowse_ld.asp?cat=4&sc="&rs("工位")&">"&rs("工位")&"</a></td>"
	response.write"<td><a href=../wokermanage/emrecord.asp?CardID="&rs("工号")&">"&rs("工号")&"</a></td>"
        response.write"<td>"&rs("姓名")&"</td>"
        response.write"<td>"&rs("职称")&"</td>"
        response.write"<td>"&rs("担当日")&"</td>"
	response.write"<td>&nbsp;&nbsp;"
       ' response.write"<a href=lcbrowse_ld.asp?act=cancel&liid="&rs("lineid")&">缩位</a>&nbsp;&nbsp;"
       ' response.write"<a href=LcEdit.asp?lineID="&rs("lineid")&">原作业者变更</a>&nbsp;&nbsp;"
       ' if hour(now()) < 10 then   
       ' response.write"<a href=LcAdd.asp?lineID="&rs("lineid")&">替位安排</a>&nbsp;&nbsp;"
      '  else
       ' response.write"老板,超10点了"
       ' end if
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&page+1&">下页<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&pagecount&">尾页<a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=Lcbrowse_ld.asp?page=1>首页<a>&nbsp;&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&page-1&">上页<a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=Lcbrowse_ld.asp?page=1>首页<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&page-1&">上页<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&page+1&">下页<a>&nbsp;"
Response.Write "<a href=Lcbrowse_ld.asp?page="&pagecount&">尾页<a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"个工位"
end sub


%>