<!--#include file="includes/db.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�û�ע��</title>
<script language="JavaScript" type="text/javascript">
   function validate(){ 
     var Username 	= document.FormRegister.Username.value;
     var Password	= document.FormRegister.Password.value;
     var ConfirmPWD     = document.FormRegister.ConfirmPWD.value;
     var TrueName	= document.FormRegister.TrueName.value;
     if (Username==""){
       alert("���������û�����");
       return false;
     }
     if (Password==""){
       alert("���������룡");
       return false;
     }
     if (Password!=ConfirmPWD){
       alert("������������벻��ͬ");
       return false;
     }
   if (TrueName==""){
       alert("��������ʵ����������Ҳ����.");
       return false;
     }
   }
</script>
<link href="../../Includes/Style.css" rel="stylesheet" type="text/css">
</head>

<body>
<form  name="FormRegister" method="post" action="register.asp" onSubmit="javaScript: return validate();">
<table border="0" cellspacing="1" width="220" cellpadding="0" align="center" bgcolor="#99CCFF">
      <tr bgcolor="#99CCFF" align="center"> 
      <td height="35" colspan="2"><b>��&nbsp;��&nbsp;ע&nbsp;��</b></td>
      </tr>
      <tr bgcolor="#FFFFFF" align="center"> 
      <td width="68" height="25">�˺�</td>
      <td width="152" height="25">
      <input type="text" name="Username" size="20" class="smallinput"></td>
      </tr>
      <tr bgcolor="#FFFFFF" align="center"> 
      <td height="25">����</td>
      <td height="25">
      <input type="password" name="Password" size="20" class="smallinput"></td>
      </tr>
      <tr bgcolor="#FFFFFF" align="center"> 
      <td height="25">ȷ������</td>
      <td height="25"><input type="password" name="ConfirmPWD" size="20" class="smallinput"></td>
      </tr>
      <tr bgcolor="#FFFFFF" align="center"> 
      <td height="25">��ʵ����</td>
      <td height="25">
      <input name="TrueName" size="20" class="smallinput"></td>
      </tr>
      <tr bgcolor="#FFFFFF">
      	<td colspan="2" align="center">
        <input name="Enter" type="submit" class="smallInput" id="Enter" value="ȷ ��">
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Clear" type="reset" class="smallInput" id="Clear" value="�� ��"> 
      </td>
      </tr>
</table>
</form>
</body>
</html>

<%
'����ǰȷ��'	
if Request.form("Enter") =  "ȷ ��"  then
   	call openDB()	
        sql="select * from t_fmea_user where username ='"&request("Username")&"' or truename='"&request("truename")&"'"
        rs.open sql,conn,1,1
        If rs.eof then
	sql="INSERT INTO t_fmea_user (Username, Pwd, TrueName,Power) " _
	    &"VALUES( '"&request("Username")&"', '"&request("Password")&"', '"&request("TrueName")&"' , 0)"	
	conn.execute(sql)
	call closeDB() 
	response.Redirect "login.asp"	  
        Else
        response.write "<script language='JavaScript'>alert('���д��û�!');</script>"  
        response.end
        End if    
end if	
%>