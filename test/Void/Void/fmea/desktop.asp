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
-->
</style>
</head>

<body  bgcolor="#f9f9f9">
<table width ="100%" border ="0">
<tr><td width="50%" valign="top">

  <table width="100%" border="0">
  <tr><td><span class="STYLE1">
  <% = request.cookies("fmea.truename")  %> ��ӭ����빤��FMEA����ϵͳ</span>
  </td></tr>
  <tr><td>&nbsp;</td></tr>
  <tr><td><span class="b1">ϵͳ˵����</span></td></tr>
  <tr><td><span class="b1">&nbsp;&nbsp;��˾�ڵĻ��ֽ϶࣬�������ֵ�FMEA�������ɲ�ͬ�ĵ�����������ʵʩ��������ʱ��FMEA����EXCEL�ļ��������ɣ��������໥�ο�������ȷ�ϣ��������ɸ�ϵͳ��</span></td></tr>
  <tr><td><span class="b1">&nbsp;</td></tr>
  <tr><td><span class="b1">��ѡ����Ĳ�����</td></tr>
   <% 
   if request.cookies("fmea.truename") = "" then
      response.write "<tr><td><span class='b1'>&nbsp;1.<a href ='Login.asp' target ='_parent'>��¼</a></span></td></tr>"
   else
      response.write "<tr><td><span class='b1'>&nbsp;1.<a href ='PjManage/Pj_Add.asp'>��¼�»���FMEA</a></span></td></tr>"
   end if 
   %>
  <tr><td><span class="b1">&nbsp;2.����,���±�ѡ��ϵ��</span></td></tr>
  <tr><td>
     <table width="100%" border="0" cellpadding="1" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" align="center">
     <% call MainBody() %>
     </table>
  </td></tr> 
  </table>

</td><td valign="top">

  <table width="100%" border="0" bgcolor="#f0f0f0" cellpadding="1" cellspacing="1">
  <tr><td  width = "12%"><span class="b1">��������:</span></td></tr> 
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
