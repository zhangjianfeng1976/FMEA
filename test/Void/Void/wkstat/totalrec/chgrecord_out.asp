<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "查询" 	                
            select case Request("cata")
              case "高级查询"
              case "模糊查询"
	         Response.Redirect "chgrecord_out.asp?Cat=0&Sc="&Request("Search")			   		
          end select            
End Select


pgs =2500
select case request("cat")
       case "" 
       sql="SELECT top 300 * FROM v_wdairy_out order by 日期 DESC"
       pgs =20
       case 0
       sql="SELECT  * FROM v_wdairy_out where 原作业者工号 like '%"&request("Sc")& _
           "%' Or 替位者工号 like '%"&request("Sc")& _
           "%' Or LINE like '%"&request("Sc")& _
           "%' Or 机种 like '%"&request("Sc")& _
           "%' Or 工程 like '%"&request("Sc")& _
           "%' Or 工位 like '%"&request("Sc")& _
           "%' ORDER BY 日期"                      
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
<link href="Includes/Style.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<body>
<form action="chgrecord_out.asp">
<table width="100%" border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr>
<td bgcolor="#e6f7ff" colspan ='11'>
<table width="100%"><tr><td align='left'>
<b><font face="华文新魏" size="3" color="#CC6600">工程外人员替位记录</font></b>
</td><td align='right'>
<select name='cata' class="smallinput">
<option>模糊查询</option>
<option>按日期查询</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="查询" name="OK" class="smallInput">
</td></tr>
</table>
</td>
</tr> 

<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="20">日期</td>
<td width="8%">Line</td>
<td width="8%">机种</td>
<td width="8%">工程</td>
<td width="8%">工位</td>
<td width="8%">原作业者工号</td>
<td width="8%">姓名</td>
<td width="8%">原由</td>
<td width="8%">替位者工号</td>
<td width="8%">替位</td>
<td width="8%">职称</td>
</tr>

<% Call dairy %>

<tr>
<td width="75%" colspan=11 height="30" bgcolor="#EFEFEF" align='center'>
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
Sub dairy()
   if rs.eof then
      response.write "<tr><td width=75% colspan=12>"
      response.write "还没有记录。"
      response.write "</td></tr>"
   end if

      n=1 
      while not rs.eof and n<= rs.pagesize
      line = rs("line")
      response.write "<tr>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("日期")&"</td>"
      response.write "<td width='9%' height='16' bgcolor='#FCFCFC'>" 
      response.write line
      response.write "</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("机种")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("工程")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("工位")&"</td>" 
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("原作业者工号")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("姓名")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("原由")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("替位者工号")&"</td>"
      response.write "<td bgcolor='#e6F6F6' align ='center'>"&rs("替位")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("职称")&"</td>"
      response.write "</tr>"
      rs.MoveNext
      n=n+1
      wend
   
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&page+1&">下页<a>&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&pagecount&">尾页<a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=chgrecord_out.asp?page=1>首页<a>&nbsp;&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&page-1&">上页<a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=chgrecord_out.asp?page=1>首页<a>&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&page-1&">上页<a>&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&page+1&">下页<a>&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&pagecount&">尾页<a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub

%>