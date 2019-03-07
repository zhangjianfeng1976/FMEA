<!--#include file="includes/db.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>首页</title>
<link href="Includes/Style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE2 {color: #0000FF}
-->
</style>
</head>

<body  bgcolor="#f9f9f9">
<table width ="100%" border ="0">
<tr><td width="50%" valign="top">

  <table width="100%" border="0">
  <tr><td><span class="STYLE1">
  <% = request.cookies("fmea.truename")  %> 欢迎你进入工程FMEA管理系统</span>
  </td></tr>
  <tr><td>&nbsp;</td></tr>
  <tr><td><span class="b1">系统说明：</span></td></tr>
  <tr><td><span class="b1">&nbsp;&nbsp;公司内的机种较多，各个机种的FMEA的评价由不同的担当进行评价实施，由于现时的FMEA是用EXCEL文件进行作成，不利于相互参考及对照确认，所以作成该系统。</span></td></tr>
  <tr><td><span class="b1">&nbsp;</td></tr>
  <tr><td><span class="b1">请选择你的操作：</td></tr>
   <% 
   if request.cookies("fmea.truename") = "" then
      response.write "<tr><td><span class='b1'>&nbsp;1.<a href ='Login.asp' target ='_parent'>登录</a></span></td></tr>"
   else
      response.write "<tr><td><span class='b1'>&nbsp;1.<a href ='PjManage/Pj_Add.asp'>登录新机种FMEA</a></span></td></tr>"
   end if 
   %>
  <tr><td><span class="b1">&nbsp;2.或者,从下表选择系列</span></td></tr>
  <tr><td>
     <table width="100%" border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" align="center">
     <% call MainBody() %>
     </table>
  </td></tr> 
  </table>

</td><td valign="top">

  <table width="100%" border="0" bgcolor="#f0f0f0" cellpadding="1" cellspacing="1">
  <tr><td  width = "12%"><span class="b1">更新履历:</span></td></tr> 
  <tr><td>2017-1-24</td><td>[主界面]主控、逻辑完成</td></tr>
  </table>

</td></tr>
</table>

</body>
</html>

<%
Sub MainBody()
sql="select distinct serial from t_fmea_pj"
call openDB()
rs.open sql,conn,1,1
	
if rs.eof then
      response.write "<tr><td width=75% colspan=3>"
      response.write "</td></tr>"
else 
      for i = 1 to rs.recordcount/8 + 1
	   response.write "<tr>"
      for aa = 1 to 8
	    if rs.eof then
               response.write "<td bgcolor='#efefef'></td>"
	    else	  
	       response.write "<td width='12%' height='20' bgcolor='#FCFCFC'>"        
               response.write "<a href=PjManage/Pj_Browse.asp?cat=1&Sc="&rs("Serial")&"><font color='#0000FF'>"
	       response.write rs("Serial")
               response.write "</font></a>"    
               rs.MoveNext 
	       response.write "</td>"
	     end if 
	   next
	   response.write "</tr>" 
	  next      	   	  
end if
call closeDB()
End Sub
%>
