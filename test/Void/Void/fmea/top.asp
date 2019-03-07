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

<script type="text/javascript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>

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

#pos_yc {
	position:absolute;width:20px;height:20px;z-index:1;
}

#pos_hj {
	position:absolute;width:20px;height:20px;z-index:1;
}

#pos_an {
	position:absolute;width:20px;height:20px;z-index:1;
}

#Layer_yc {	
    position:absolute; padding:0px;width:800px;height:500px;z-index:1;left: -7px;top: 17px;visibility: hidden;
}

#Layer_hj {	
    position:absolute;padding:0px;width:500px;height:500px;z-index:1;left: -7px;top: 17px;visibility: hidden;
}	

#Layer_an {	
    position:absolute;padding:0px;width:500px;height:500px;z-index:1;left: -7px;top: 17px;visibility: hidden;
}

-->
</style>


</head>

<body   bgcolor="#0099CC">
<!--<body   bgcolor="#6699cc" onLoad="popwin('memo/20160322.html',300,350)"> 
<body marginheight="0" topmargin="0" bottommargin="0" leftmargin="0" bgcolor="#f9f9f9">-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr><td>
<a href ="default.asp" target ="_parent">
<img border="0" src="images/kyocera.gif" width="100" height="30" align='left'>
</a>
</td><td align="center">
<font face="华文中宋" size="5" color="#ffffff">工程FMEA管理系统</font>
</td>
<td align ="right" valign="top">
<div style ="float:right ;">
<div style ="float:right ;">&nbsp;<font size="2" color="#ffffff">Ver.2017-01-24</font></div>
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
</td></tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3" >
<tr height ="20">
<td width="5%" align="center" bgcolor="#55CCFF"
    onmouseover="MM_showHideLayers('Layer_yc','','show')" 
    onmouseout="MM_showHideLayers('Layer_yc','','hide')">
<div id="pos_yc"><div id="Layer_yc" 
     onMouseOver="MM_showHideLayers('Layer_yc','','show')" 
     onMouseOut="MM_showHideLayers('Layer_yc','','hide')">
<table width="160" border="0" align="left" cellpadding="0" cellspacing="1">
<tr>
<td height = '20' width ='80' align="center" bgcolor="#99CCFF">
<span class="STYLE1"><a style="COLOR: #ffffff" href="PjManage\Pj_Browse.asp" target="MainWindow">机种一览</a></span></td> 
<td height = '20' width ='80'  align="center" bgcolor="#99CCFF">
<span class="STYLE1"><a style="COLOR: #ffffff" href="PjManage\Appr_Browse.asp" target="MainWindow">承认流程</a></span></td>   

</tr>
</table>
  </div>
</div>
<span class="style1"><a style="COLOR: #ffffff">机种管理</a></span></td>

<td width="5%"  align="center" bgcolor="#55CCFF"
    onmouseover="MM_showHideLayers('Layer_hj','','show')" 
    onmouseout="MM_showHideLayers('Layer_hj','','hide')">
<div id="pos_hj"><div id="Layer_hj" 
     onMouseOver="MM_showHideLayers('Layer_hj','','show')" 
     onMouseOut="MM_showHideLayers('Layer_hj','','hide')">
<table width="160"  border="0" align="left" cellpadding="0" cellspacing="1">
<tr>
<td height = '20' align="center" bgcolor="#99CCFF">
<span class="STYLE1"><a style="COLOR: #ffffff" href="ProcessMan\Process_Browse.asp" target="MainWindow">作业管理</a></span></td>
<td height = '20' align="center" bgcolor="#99CCFF"> 
<span class="STYLE1"><a style="COLOR: #ffffff" href="ProcessMan\Advice_Browse.asp" target="MainWindow">措施管理</a></span></td>
</tr>
</table>
  </div>
</div>
<span class="style1">评价管理</span></td>

<td width="5%"  align="center" bgcolor="#55CCFF"
    onmouseover="MM_showHideLayers('Layer_an','','show')" 
    onmouseout="MM_showHideLayers('Layer_an','','hide')">
<div id="pos_an"><div id="Layer_an" 
     onMouseOver="MM_showHideLayers('Layer_an','','show')" 
     onMouseOut="MM_showHideLayers('Layer_an','','hide')">
<table width="240"  border="0" align="left" cellpadding="0" cellspacing="1">
<tr>
<td height = '20' align="center" bgcolor="#99CCFF">
<span class="STYLE1"><a style="COLOR: #ffffff" href="" target="MainWindow"></a></span></td>
<td height = '20' align="center" bgcolor="#99CCFF">
<span class="STYLE1"><a style="COLOR: #ffffff" href="" target="MainWindow"></a></span></td>
<td height = '20' align="center" bgcolor="#99CCFF"> 
<span class="STYLE1"><a style="COLOR: #ffffff" href="" target="MainWindow"></a></span></td>
</tr>
</table>
  </div>
</div>
<span class="style1">分析统计</span></td>


<td width="60%" align="right">
<a style="COLOR: #ffffff;" href="Default.asp" target="_top"></a>&nbsp;&nbsp;</td>
</tr>

</table>

 </body>
</html>



