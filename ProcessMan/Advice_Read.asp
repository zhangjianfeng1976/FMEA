<!--#include file="../includes/db.asp"-->
<%
pgs =2500
select case request("cat")
       case "" 
         ttl = request.cookies("fmea.Pjkey")
         sql="SELECT * FROM v_Fmea_total where Pjkey = '"&request.cookies("fmea.Pjkey")&"' order by SubPjName,PageNo,WorkQueue"
         pgs = 10
       case 0
          ttl = request.cookies("fmea.Pjkey")
          sql = "SELECT * FROM v_Fmea_total where "
          sql = sql & "(AdvContent like '%" & request("sc") & "%' or "
          sql = sql & "Rspnser like '%" & request("sc") & "%' or "
          sql = sql & "ActionContent like '%" & request("sc") & "%'"
          sql = sql & ") and Pjkey = '"&request.cookies("fmea.Pjkey")&"' and  AdID is not null"
       case 1
          ttl = "全部"
          sql = "SELECT * FROM v_Fmea_total where "
          sql = sql & "(AdvContent like '%" & request("sc") & "%' or "
          sql = sql & "Rspnser like '%" & request("sc") & "%' or "
          sql = sql & "ActionContent like '%" & request("sc") & "%'"
          sql = sql & ") order by ItmID"
       case 2
         ttl = request.cookies("fmea.Pjkey")
         sql="SELECT * FROM v_Fmea_total WHERE ItmID = "& request("ItmID")  
       case 3
         ttl = request.cookies("fmea.Pjkey")
         sql="SELECT * FROM v_Fmea_total WHERE Pjkey = '"&request.cookies("fmea.Pjkey")&"' and  AdID is not null order by SubPjName,PageNo,WorkQueue"     
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

<title>改善措施</title>
</head>

<body bgcolor ="#FCFCFC">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF"><td height="30" align="center" colspan=14><b><% = ttl %>改善措施</b></td></tr>
<tr bgcolor="#FCFCFC" align="center">
<td width="2%"  rowspan ="2">顺序</td>
<td width="8%" height="20" rowspan ="2">工程名称</td>
<td width="4%"  rowspan ="2">标准书页码</td>
<td width="8%"  rowspan ="2">部品编号</td>
<td width="15%" rowspan ="2">建议措施</td>
<td width="5%"  rowspan ="2">责任者</td>
<td width="8%"  rowspan ="2">目标完成日</td>
<td width="18%" colspan ="6">措施结果</td>
</tr><tr bgcolor="#FCFCFC" align="center">
<td width="10%">采用措施</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RI</td>
<td width="5%"></td>
</tr>

<% 


call MainBody()

 %>


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
    response.write"<td>"&rs("WorkQueue")&"</td>"
    response.write"<td>"&rs("SubPjName")&"</td>"	
    response.write"<td>"&rs("PageNo")&"</td>"
    response.write"<td>"&rs("PartsNo")&"</td>"	
    response.write"<td>"&rs("AdvContent")&"</td>"	
	response.write"<td>"&rs("Rspnser")&"</td>"
    response.write"<td>"&rs("FshDate")&"</td>"
    response.write"<td>"&rs("ActionContent")&"</td>"
    response.write"<td>"&rs("ResultS")&"</td>"
    response.write"<td>"&rs("ResultO")&"</td>"
    response.write"<td>"&rs("ResultD")&"</td>"
    response.write"<td>"&rs("ResultRPN")&"</td>" 
    response.write"<td>"&rs("WorkStation")&"</td>" 
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
End Sub

%>

