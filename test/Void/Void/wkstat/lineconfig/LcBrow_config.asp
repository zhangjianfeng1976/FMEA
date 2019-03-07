<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "查询" 	                
            select case Request("cata")
              case "高级查询"
              case "模糊查询"
	            Response.Redirect "LcBrow_config.asp?Cat=0&Sc="&Request("Search")			   		
              case "无效担当"
                    Response.Redirect "LcBrow_config.asp?Cat=1&Sc="&Request("Search")
              case "己变担当"
                    Response.Redirect "LcBrow_config.asp?Cat=2&Sc="&Request("Search")
              case "按场所查询"
                    Response.Redirect "LcBrow_config.asp?Cat=3&Sc="&Request("Search")
              case "按机番查询"
                    Response.Redirect "LcBrow_config.asp?Cat=4&Sc="&Request("Search")		
          end select            
End Select


lineID=request("lineID")
Select case request("action")
       case "changeN"
            sql = "Update wt_lineconfig set wkstatus = 'N' where LineId = "&LineID
            call openDB()
            conn.execute(sql)
            call closeDB()
       case "changeY"
            sql = "Update wt_lineconfig set wkstatus = 'Y' where LineId = "&LineID
            call openDB()
            conn.execute(sql)
            call closeDB()
end select

select case request("cat")
       case "" 
       sql="SELECT * FROM  wt_lineconfig"
       pgs =25
       case 0
       sql="SELECT * FROM wt_lineconfig where lineid like '%"&request("Sc")& _
           "%' Or line like '%"&request("Sc")& _
           "%' Or model like '%"&request("Sc")& _
           "%' Or wogroup like '%"&request("Sc")& _
           "%' Or workstation like '%"&request("Sc")& _
           "%' Or originworker like '%"&request("Sc")& _
           "%' ORDER BY lineid" 
       pgs =2500      
       case 1
       sql="SELECT * FROM  wt_lineconfig where Wkstatus ='N'"
       pgs =2500
       case 2
       sql="SELECT * FROM  wt_lineconfig where changedate is not null"
       pgs =2500       
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
<form action="LcBrow_config.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr><td height="20" align='right' bgcolor="#e6f7ff">
<select name='cata' class="smallinput">
<option>模糊查询</option>
<option>无效担当</option>
<option>己变担当</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="查询" name="OK" class="smallInput">
</td>
</tr>
</table>
</form>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=9>
<b>工程配置</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="left">
<td width="100%" height="20" colspan = 9>
*下记【有效】列指的是记录的状态是否有效，如果是Y的话，代表该员工仍担当此工位，如果是N的话，表示该记录己无效，工位已被另一员工替代
</td></tr>
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="20">Line</td>
<td width="8%">机种</td>
<td width="10%">工程</td>
<td width="10%">工位</td>
<td width="10%">原担当者</td>
<td width="4%">有效</td>
<td width="8%">担当日期</td>
<!--<td width="20%">操作</td>-->
</tr>

<% call MainBody() %>

<tr>
<td width="75%" colspan=9 height="30" bgcolor="#EFEFEF" align='center'>
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
	response.write "<tr><td colspan=6>"
	response.write "没有记录。"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
	response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<td>"&rs("Line")&"</td>"	
        response.write"<td>"&rs("Model")&"</td>"
        response.write"<td>"&rs("WoGroup")&"</td>"
        response.write"<td>"&rs("WorkStation")&"</td>"
	response.write"<td>"&rs("Originworker")&"</td>"
        response.write"<td>"&rs("Wkstatus")&"</td>"
        response.write"<td>"&rs("changedate")&"</td>"
'       response.write"<td><a href=LcEdit.asp?lineID="&rs("lineid")&">变更原担当者</a>&nbsp;&nbsp;"
'       response.write"<a href=LcBrow_config.asp?action=changeN&lineID="&rs("lineid")&">转为无效</a>&nbsp;&nbsp;"
'       response.write"<a href=LcBrow_config.asp?action=changeY&lineID="&rs("lineid")&">转为有效</a>&nbsp;&nbsp;"
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&pagecount&">尾页</a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=LcBrow_config.asp?page=1>首页</a>&nbsp;&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&page-1&">上页</a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=LcBrow_config.asp?page=1>首页<a>&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&page-1&">上页</a>&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=LcBrow_config.asp?page="&pagecount&">尾页</a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub


%>
