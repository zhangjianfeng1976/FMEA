<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "查询" 	                
            select case Request("cata")
              case "高级查询"
              case "模糊查询"
	           Response.Redirect "Browse.asp?Cat=0&Sc="&Request("Search")			   		
              case "工程外人员"
                    Response.Redirect "Browse.asp?Cat=1&Sc="&Request("Search")
          end select            
       Case "人员分配工程工位输入"
             Response.Redirect "add.asp"
End Select


ID=request("ID")
'Select case request("Action")
'       case "Delete" 	
'            sql="Delete wt_employbase Where ID ="&ID       
'            call openDB()
'            conn.execute(sql)
'            call closeDB()
'end select

pgs =2500
select case request("cat")
       case "" 
       sql="SELECT * FROM  wt_employbase where statuss='未离职'"
       pgs =25
       case 0
       sql="SELECT * FROM wt_employbase where (cardid like '%"&request("Sc")& _
           "%' Or name like '%"&request("Sc")& _
           "%' Or vocation like '%"&request("Sc")& _
           "%' Or dpcode like '%"&request("Sc")& _
           "%' Or department like '%"&request("Sc")& _
           "%' Or sex like '%"&request("Sc")& _
           "%') ORDER BY cardid"    
       case 1
        sql ="select * from wt_employbase where cardid not in (select originworker from wt_lineconfig where wkstatus = 'y')"  
       case 2
       sql="SELECT * FROM wt_employbase where dpcode like '%"&request("Sc")&"%' and  statuss='未离职'  ORDER BY cardid"         
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
<script language="javascript" type="text/javascript">
<!--
function jump() 
{
 if(event.keyCode==13)event.keyCode=9
}
-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>工位人员管理</title>
</head>

<body>
<form action="Browse.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr><td height="20" align='right' bgcolor="#e6f7ff">
<select name='cata' class="smallinput">
<option>模糊查询</option>
<option>工程外人员</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="查询" name="OK" class="smallInput">
<a href="../includes/LineConfigBatch.xls">人员工位输入</a>
<!--<input type='submit' value='人员分配工程工位输入' name='OK' class='smallInput'>-->

</td>
</tr>
</table>
</form>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=11>
<b>人员列表</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="20">工号</td>
<td width="6%">姓名</td>
<td width="6%">职称</td>
<td width="3%">性别</td>
<td width="4%">系编号</td>
<td width="30%">所属</td>
<td width="8%">入厂日期</td>
<td width="8%">离职日</td>
<td width="8%">状态</td>
<td width="8%">操作</td>
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
	response.write "<tr><td colspan=6>"
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
        response.write"<td>"&rs("DPcode")&"</td>"
        response.write"<td>"&rs("department")&"</td>"
	response.write"<td>"&rs("entrancedate")&"</td>"
        response.write"<td>"&rs("lastwd")&"</td>"
        response.write"<td>"&rs("statuss")&"</td>"
        response.write"<td><a href=emEdit.asp?ID="&rs("id")&">"&"报考勤</a>&nbsp;&nbsp"
        'response.write"<a href=add.asp?CardID="&rs("CardID")&">"&"安排入位</a>&nbsp;&nbsp;"
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=browse.asp?page="&pagecount&">尾页</a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=browse.asp?page=1>首页</a>&nbsp;&nbsp;"
Response.Write "<a href=browse.asp?page="&page-1&">上页</a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=browse.asp?page=1>首页</a>&nbsp;"
Response.Write "<a href=browse.asp?page="&page-1&">上页</a>&nbsp;"
Response.Write "<a href=browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=browse.asp?page="&pagecount&">尾页</a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub


%>
