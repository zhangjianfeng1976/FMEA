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
.zw1{font-size: 9pt; font-family:"Verdana", "Arial", "宋体"; color: #0000FF}
.zw2{font-size: 9pt; font-family:"Verdana", "Arial", "宋体"; color: #000000}
.ena{font-size: 9pt; font-family:"Verdana", "Arial", "宋体"; color: #aaaaaa}
-->
</style>
</head>

<body  bgcolor="#FCFCFC">
<table width ="100%" border ="0">
<tr><td width="50%" valign="top">

  <table width="100%" border="0">
  <tr><td><span class="STYLE1">
  <% = request.cookies("fmea.truename")  %> 欢迎你进入工程FMEA管理系统</span>
  </td></tr>
  

  
  <tr><td>&nbsp;</td></tr>
 
  <tr><td><span class="b1">请选择你的操作：</td></tr>
   <tr><td>&nbsp;1. 分析页面</td></tr> 
   <tr><td>
	   <table width="100%" border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" align="center">
	   <tr><td width ="20%" height = "20" bgcolor="#efefef"><a href ="analyze/Serial.asp">机种部品别分析</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="analyze/author.asp">团队部品别分析</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="analyze/Serial_workcata.asp">机种作业别分析</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="analyze/Author_workcata.asp">团队作业别分析</a></td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
	   </tr> 

	   <tr><td width ="20%" height = "20" bgcolor="#efefef"><a href ="analyze/Serial_RI_CNT.asp">机种评分表</a></td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
	   </tr>
	   </table> 
	</td></tr>
	<tr><td>&nbsp;2.流程状态 </td></tr> 
	<tr><td>
	   <table width="100%" border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" align="center">
	   <tr>
	    <%
	      if request.cookies("fmea.power") > "1" then 
	         response.write "<td width ='20%' height = '20' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=1>作成中项目</a></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=2>待审查项目</a></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=3>待承认项目</a></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=4>待配布项目</a></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=5>完成配布项目</a></td>"
		  else 
		     response.write "<td width ='20%' height ='20' bgcolor='#efefef'><span class='ena'>作成中项目</span></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><span class='ena'>待审查项目</span></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><span class='ena'>待承认项目</span></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><span class='ena'>待配布项目</span></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><span class='ena'>完成配布项目</span></td>"
          end if		  
		%>   
	   </tr> 
	   </table> 
  </td></tr>
  <tr><td>&nbsp;3.过去问题，相似部品及评价标准化 </td></tr> 
   <tr><td>
	   <table width="100%" border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" align="center">
	   <tr>
		   
		   <td width ="20%" height = "20" bgcolor="#efefef"><a href ="His_Question/PartsFamily.asp">相似部品</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="His_Question/FailBase.asp">过去问题库</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="His_Question/SafeParts.asp">安全重要部品</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="includes/FMEA评分标准化指南.pdf"  target ="_blank">KDTCN FMEA评价标准</a></td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
	   </tr> 
	   </table> 
  </td></tr>

   <% 
   if request.cookies("fmea.truename") = ""  then
	  response.write "<tr><td>&nbsp;4.[<a href ='Login.asp' target ='_parent'><span class='zw2'>登录</span></a>]</td></tr>"
	  response.write "<tr><td>&nbsp;5.<span class='zw2'>或者,从下表选择系列浏览</span></td></tr>"
   else
      if request.cookies("fmea.power") > "1" then 
	     response.write "<tr><td>&nbsp;4.<a href ='PjManage/Pj_Add.asp'><span class='zw1'>登录新机种FMEA</span></a></td></tr>"
	  else 
	     response.write "<tr><td>&nbsp;4.<span class='zw1'>登录新机种FMEA</span></td></tr>"
	  end if 
	  response.write "<tr><td>&nbsp;5.<span class='zw1'>或者,从下表选择系列操作</span></td></tr>"
   end if 
   %>
   
	<tr><td>
		 <table width="100%" border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" align="center">
		 <% call MainBody() %>
		 </table>
		
  </td></tr>
  
  </table>

</td><td valign="top">
		<table width="100%" border="0" bgcolor="#f0f0f0" cellpadding="1" cellspacing="1">
		<tr><td colspan ="2"><span class="b1">系统说明：</span></td></tr>
		<tr><td colspan ="2"><span class="b1">
		<div>&nbsp;&nbsp;&nbsp;&nbsp;公司内的机种较多，各个机种的FMEA的评价由不同的担当进行评价实施，由于现时的FMEA是用EXCEL文件进行作成，不利于相互参考及对照确认，所以作成该系统。</div>
		<div>&nbsp;&nbsp;&nbsp;&nbsp;开发过程中遇到的课题较多，将会逐渐完善。</div>
		</span></td></tr>
		<tr><td colspan ="2"><img border="0" src="images/fmea4.jpg" width="416" height="300" align='left'></td></tr>

		<tr><td><span class="b1">&nbsp;</td></tr>
		<tr><td  width = "12%"><span class="b1">更新履历:</span></td></tr> 
		<tr><td>2018-07-24</td><td><span class="STYLE1">[主界面]追加安全重要部品</span></td></tr>
		<tr><td>2017-12-26</td><td>[主界面]追加KDTCN FMEA评价标准</span></td></tr>
		<tr><td>2017-6-28</td><td>[评价页面]过去问题的追加</td></tr>
		<tr><td>2017-3-10</td><td>[评价及对应页面]优化显示及对应输入</td></tr>
		<tr><td>2017-2-10</td><td>[输入页面]10点法变更为4点法</td></tr>
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
	  for i = 1 to rs.recordcount/6 + 1
	   response.write "<tr>"
	  for aa = 1 to 6
	    if rs.eof then
			   response.write "<td bgcolor='#f9f9f9'></td>"
	    else	  
	       response.write "<td width='12%' height='20' bgcolor='#efefef'>"        
			   response.write "<a href=PjManage/Pj_Browse.asp?serial="&trim(rs("Serial"))&"><font color='#0000FF'>"
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
