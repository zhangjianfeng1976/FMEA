<!--#include file="includes/db.asp"-->
<html>
<head>
<title>工程FMEA管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="includes/Style.css">
<script language="JavaScript" type="text/javascript"><!--
   function validate(){ 
     var Username 	= document.FormLogin.Username.value;
     var Password	= document.FormLogin.Password.value;
     if (Username==""){
       alert("请您输入用户名！");
       return false;
     }
     if (Password==""){
       alert("请输入密码！");
       return false;
     }
   }
-->
</script>

<style type="text/css">
<!--
.style1  {font-size:18px; color:#ffffff}
.style2 {font-size: 12px}
body {
	background-image: url(images/1.jpg);
}

.STYLE3 {
	font-size: 12pt;
	color: #0033FF;
}
-->
</style>
</head>

<body>
<form name="FormLogin" method="POST" action="login.asp" onSubmit="javaScript: return validate();">
<table align="center" border="0" cellspacing="0" width="200"  cellpadding="1" bgcolor="#6699CC">
      <tr><td bgcolor="#6699CC" colspan ="2">
      <p align="center" class="style1">系统登陆</td>
      </tr>
      <tr> 
      <td width="64" height="17" bgcolor="#FCFCFC" align="right"><span class="style2">&#36134;&#21495;</span></td>
      <td width="136" height="17" bgcolor="#FCFCFC" align="left"> &nbsp;<input type="text" name="Username" size="12" class="smallinput" >
      </td></tr>
      <tr><td  bgcolor="#FCFCFC" align="right"><span class="style2">&#23494;&#30721;</span></td>
      <td  bgcolor="#FCFCFC" align="left"> &nbsp;<input type="password" name="Password" size="12" class="smallinput">
      </td></tr>
      <tr align="center"><td colspan="2" height="10"  bgcolor="#FCFCFC">
      <input name="Enter" type="submit" class="smallInput" id="Enter" value="&#36827;&#20837;&#31995;&#32479;">&nbsp;
      <span class="style2">[<a href="Register.asp">注册</a>] </span>
      </td></tr> 	 
  </table>
</form>
</body>
</html>


<%
Dim Username
Dim Password 
	
If Request.form("Enter") =  "进入系统"  then
   Username=Request("Username")
   Password=Request("Password")
	
   sql="SELECT * FROM t_fmea_user WHERE Username= '" + Username +"' AND Pwd = '" +Password +"'"
   Call OpenDB()
   Set result=conn.execute(sql)
	
   If not result.eof then   
      response.cookies("fmea.username")=Result("username")
      response.cookies("fmea.truename")=Result("truename")
      'Session("Truename")=Result("truename")
      response.cookies("fmea.power")=Result("power")
      Response.Redirect "index.asp"
   Else
      Response.Write "<script language='JavaScript' type='text/javascript'>alert('有错误发生！请确认用户名及密码正确');</script>"		
   End if
   CloseDB()
End if
%>
