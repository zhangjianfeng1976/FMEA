<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>工程FMEA管理系统</title>

<link href="Includes/Style.css" rel="stylesheet" type="text/css">
<Script LANGUAGE = "JAVASCRIPT" type="text/javascript">
<!--Java Script Start
var winconf = "menubar=no,toolbar=no,location=no,directories=no,status=no,scrollbar=no,resizeable=no,";
function popwin(filename,w,h) 
{
a=window.screen.height;
b=window.screen.width;
a=(a-h)/4 ;
b=(b-w)/2 ;
popwin = window.open(filename,"Copyright",winconf+"left="+ 0 +" top="+ 0 +" width="+w+",height="+h);
popwin.focus();
}
//Java Script End-->
</SCRIPT> 

<style type="text/css">
<!--
.style1  {font-size: 18px; color:#ffffff}
.style2  {font-size: 12px}
.style3  {font-size: 12pt; color: #0033ff;}
-->
</style>

<style type="text/css">
<!--
.title{
font-family: "华文行楷", "华文新魏";font-size: 20px; color: #FF0000;
}
.style1 {
font-size: 12px;color: #FFFFFF;
}

-->
</style>


</head>

<body bgcolor="#0099CC">

<!--
<body   bgcolor="#6699cc" onLoad="popwin('memo/20160322.html',300,350)"> 
<body marginheight="0" topmargin="0" bottommargin="0" leftmargin="0" bgcolor="#f9f9f9">
-->

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr><td>
<a href ="desktop.asp" target ="MainWindow">
<img border="0" src="images/kyocera.gif" width="100" height="30" align='left'>
</a>
</td><td align="center">
<font face="华文中宋" size="5" color="#ffffff">工程FMEA管理系统</font>
</td>
<td align ="right" valign="top">
<div style ="float:right ;">
<div style ="float:right ;">&nbsp;<font size="2" color="#ffffff">Ver.2018-07-24</font></div>
<div style ="float:right ;">&nbsp;<a href ="register.asp" target = "_parent"><font size="2" color="#ffffff">注册</font></a></div>
<div style ="float:right ;">&nbsp;

<%
if request.cookies("fmea.truename") =""  then 
   response.write "<a href = 'login.asp' target = '_parent'><font size='2' color='#ffffff'>登录</font></a>"
else
   fmea_truename = request.cookies("fmea.truename")
   fmea_power    = request.cookies("fmea.power")
   response.write "<font size='2' color='#ffffff'>"&fmea_power&"级用户:"&fmea_truename&"</font>"
   response.write "<a href ='default.asp' target ='_parent'><font size='2' color='#ffffff'>退出</font></a>"
end if
%>

</div>
</div>
<div style ="float:right ;">
<font size='2' color='#ffffff'>

</font>
</div>
</td></tr>
</table>
 </body>
</html>



