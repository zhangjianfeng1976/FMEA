<!--#include file="includes/db.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ҳ</title>
<link href="Includes/Style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE2 {color: #0000FF}
.zw1{font-size: 9pt; font-family:"Verdana", "Arial", "����"; color: #0000FF}
.zw2{font-size: 9pt; font-family:"Verdana", "Arial", "����"; color: #000000}
.ena{font-size: 9pt; font-family:"Verdana", "Arial", "����"; color: #aaaaaa}
-->
</style>
</head>

<body  bgcolor="#FCFCFC">
<table width ="100%" border ="0">
<tr><td width="50%" valign="top">

  <table width="100%" border="0">
  <tr><td><span class="STYLE1">
  <% = request.cookies("fmea.truename")  %> ��ӭ����빤��FMEA����ϵͳ</span>
  </td></tr>
  

  
  <tr><td>&nbsp;</td></tr>
 
  <tr><td><span class="b1">��ѡ����Ĳ�����</td></tr>
   <tr><td>&nbsp;1. ����ҳ��</td></tr> 
   <tr><td>
	   <table width="100%" border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" align="center">
	   <tr><td width ="20%" height = "20" bgcolor="#efefef"><a href ="analyze/Serial.asp">���ֲ�Ʒ�����</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="analyze/author.asp">�ŶӲ�Ʒ�����</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="analyze/Serial_workcata.asp">������ҵ�����</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="analyze/Author_workcata.asp">�Ŷ���ҵ�����</a></td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
	   </tr> 

	   <tr><td width ="20%" height = "20" bgcolor="#efefef"><a href ="analyze/Serial_RI_CNT.asp">�������ֱ�</a></td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
	   </tr>
	   </table> 
	</td></tr>
	<tr><td>&nbsp;2.����״̬ </td></tr> 
	<tr><td>
	   <table width="100%" border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" align="center">
	   <tr>
	    <%
	      if request.cookies("fmea.power") > "1" then 
	         response.write "<td width ='20%' height = '20' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=1>��������Ŀ</a></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=2>�������Ŀ</a></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=3>��������Ŀ</a></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=4>���䲼��Ŀ</a></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><a href=PjManage/Pj_Browse.asp?cat=5>����䲼��Ŀ</a></td>"
		  else 
		     response.write "<td width ='20%' height ='20' bgcolor='#efefef'><span class='ena'>��������Ŀ</span></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><span class='ena'>�������Ŀ</span></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><span class='ena'>��������Ŀ</span></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><span class='ena'>���䲼��Ŀ</span></td>"
		     response.write "<td width ='20%' bgcolor='#efefef'><span class='ena'>����䲼��Ŀ</span></td>"
          end if		  
		%>   
	   </tr> 
	   </table> 
  </td></tr>
  <tr><td>&nbsp;3.��ȥ���⣬���Ʋ�Ʒ�����۱�׼�� </td></tr> 
   <tr><td>
	   <table width="100%" border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" align="center">
	   <tr>
		   
		   <td width ="20%" height = "20" bgcolor="#efefef"><a href ="His_Question/PartsFamily.asp">���Ʋ�Ʒ</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="His_Question/FailBase.asp">��ȥ�����</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="His_Question/SafeParts.asp">��ȫ��Ҫ��Ʒ</a></td>
		   <td width ="20%" bgcolor="#efefef"><a href ="includes/FMEA���ֱ�׼��ָ��.pdf"  target ="_blank">KDTCN FMEA���۱�׼</a></td>
		   <td width ="20%" bgcolor="#f9f9f9">&nbsp;</td>
	   </tr> 
	   </table> 
  </td></tr>

   <% 
   if request.cookies("fmea.truename") = ""  then
	  response.write "<tr><td>&nbsp;4.[<a href ='Login.asp' target ='_parent'><span class='zw2'>��¼</span></a>]</td></tr>"
	  response.write "<tr><td>&nbsp;5.<span class='zw2'>����,���±�ѡ��ϵ�����</span></td></tr>"
   else
      if request.cookies("fmea.power") > "1" then 
	     response.write "<tr><td>&nbsp;4.<a href ='PjManage/Pj_Add.asp'><span class='zw1'>��¼�»���FMEA</span></a></td></tr>"
	  else 
	     response.write "<tr><td>&nbsp;4.<span class='zw1'>��¼�»���FMEA</span></td></tr>"
	  end if 
	  response.write "<tr><td>&nbsp;5.<span class='zw1'>����,���±�ѡ��ϵ�в���</span></td></tr>"
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
		<tr><td colspan ="2"><span class="b1">ϵͳ˵����</span></td></tr>
		<tr><td colspan ="2"><span class="b1">
		<div>&nbsp;&nbsp;&nbsp;&nbsp;��˾�ڵĻ��ֽ϶࣬�������ֵ�FMEA�������ɲ�ͬ�ĵ�����������ʵʩ��������ʱ��FMEA����EXCEL�ļ��������ɣ��������໥�ο�������ȷ�ϣ��������ɸ�ϵͳ��</div>
		<div>&nbsp;&nbsp;&nbsp;&nbsp;���������������Ŀ���϶࣬���������ơ�</div>
		</span></td></tr>
		<tr><td colspan ="2"><img border="0" src="images/fmea4.jpg" width="416" height="300" align='left'></td></tr>

		<tr><td><span class="b1">&nbsp;</td></tr>
		<tr><td  width = "12%"><span class="b1">��������:</span></td></tr> 
		<tr><td>2018-07-24</td><td><span class="STYLE1">[������]׷�Ӱ�ȫ��Ҫ��Ʒ</span></td></tr>
		<tr><td>2017-12-26</td><td>[������]׷��KDTCN FMEA���۱�׼</span></td></tr>
		<tr><td>2017-6-28</td><td>[����ҳ��]��ȥ�����׷��</td></tr>
		<tr><td>2017-3-10</td><td>[���ۼ���Ӧҳ��]�Ż���ʾ����Ӧ����</td></tr>
		<tr><td>2017-2-10</td><td>[����ҳ��]10�㷨���Ϊ4�㷨</td></tr>
		<tr><td>2017-1-24</td><td>[������]���ء��߼����</td></tr>
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
