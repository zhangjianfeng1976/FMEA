<!--#include file="../includes/db.asp"-->
<%
SqlStr_Q ="select ItmID from t_Fmea_Advice Where ItmID ='"&Request("ItmID")&"'"
call openDB()
Set result=conn.execute(SqlStr_Q)
If not result.eof then 
   call closeDB()
   Response.write "����Ŀ�ĸ��ƴ�ʩ�����룬�뷵�ز鿴.[<a href='javascript:history.back()'>����</a>]"
   Response.end 
End If 

Select Case Request("OK")
       Case "�ص���Ŀ�б�"
           Response.redirect "Process_Browse.asp"
       Case "�ص�δ�����б�"
           Response.redirect "Advice_BrInput.asp"   
       Case "��ʩ����"
           R_Sev = int(request("ResultS"))
           R_Occ = int(request("ResultO"))
           R_Det = int(request("ResultD"))
           R_RPN = Round((R_Sev * R_Occ * R_Det)^(1/3),1)
           sql= "Insert into t_Fmea_Advice"
           sql= sql & "(ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,ActionContent,"
           sql= sql & "ResultS,ResultO,ResultD,ResultRPN,AdvStatus,InUser,InTime) values ('" 
           sql= sql & request("ItmID")&"','" 
		   sql= sql & request("PjID")&"','"
		   sql= sql & request("Pjkey")&"','" 
           sql= sql & request("AdvContent")&"','" 
           sql= sql & request("Rspnser")&"','" 
           sql= sql & request("FshDate")&"','" 
           sql= sql & request("ActionContent")&"','" 
           sql= sql & request("ResultS")&"','"
           sql= sql & request("ResultO")&"','"
           sql= sql & request("ResultD")&"','"
           sql= sql & R_RPN &"','Y','"
           sql= sql & Request.Cookies("fmea.truename") &"','"
           sql= sql & Now() &"')" 
           call openDB()
           conn.execute(sql)
           call closeDB()
           Response.Redirect "Advice_BrInput.asp"			       
End Select
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript" type="text/javascript">
<!--
function jump() 
{
 if(event.keyCode==13)event.keyCode=9
}
-->
</script>
<title></title>
</head>

<body   marginheight="0" topmargin="10" bottommargin="0" leftmargin="0" bgcolor="#f9f9f9">
<form action="Advice_Add.asp" method="post">
<table width="750"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center">
<tr bgcolor="#6c9c6c">
<td colspan ="2" height ="30" align ="center">
<font size ="3" color ="ffffff">
<!-- (<< = request.cookies("fmea.PjID") >>)<< = request.cookies("fmea.PjKey") >>-<< = request("ItmID") >>  -->
(<% = request("PjID") %>)<% = request("PjKey") %>-<% = request("ItmID") %>
���ƴ�ʩ��������</font>
<input name="ItmID" type ="hidden" value ="<% = request("ItmID") %>"></td></tr>
<tr bgcolor="#FCfCFC">
<td width = "150">�����ȡ�Ĵ�ʩ����</td>
<td><input name="AdvContent" class="notline" size="100" tabindex ="1" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>������</td>
<td><input name="Rspnser" class="notline" size="25" tabindex ="2" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>Ŀ�������</td>
<td><input name="FshDate" class="notline" size="25" tabindex ="3" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>��ȡ��ʩ</td>
<td><input name="ActionContent" class="notline" size="100" tabindex ="4" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>��ʩ�����ض�Sev</td>
<td>
<input name="ResultS" type ="Radio" value ="1" tabindex ="5" onKeyDown="jump()"> 1
<input name="ResultS" type ="Radio" value ="2"> 2
<input name="ResultS" type ="Radio" value ="3"> 3
<input name="ResultS" type ="Radio" value ="4"> 4
</td></tr>
<tr bgcolor="#FCfCFC">
<td>��ʩ��Ƶ��Occ</td>
<td>
<input name="ResultO" type ="Radio" value ="1" tabindex ="6" onKeyDown="jump()"> 1
<input name="ResultO" type ="Radio" value ="2"> 2
<input name="ResultO" type ="Radio" value ="3"> 3
<input name="ResultO" type ="Radio" value ="4"> 4
</td></tr>
<tr bgcolor="#FCfCFC">
<td>��ʩ����̽���Det</td>
<td>
<input name="ResultD" type ="Radio" value ="1" tabindex ="7" onKeyDown="jump()"> 1
<input name="ResultD" type ="Radio" value ="2"> 2
<input name="ResultD" type ="Radio" value ="3"> 3
<input name="ResultD" type ="Radio" value ="4"> 4
</td>
</tr>
<tr bgcolor="#FCfCFC" height ="30">
<td colspan = "2" align ="center" >
<input name="OK" value="��ʩ����" type="submit" class="smallInput">
<% 
if request("cat") = 1 then 
else
response.write "<input name='OK' value='�ص���Ŀ�б�' type='submit' class='smallInput'>"
response.write "<input name='OK' value='�ص�δ�����б�' type='submit' class='smallInput'>"
end if
%>
</td></tr>
</table>
</form>
</body>
</html>

