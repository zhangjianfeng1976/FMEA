<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK")  
       Case "��¼"
           if request("Serial") ="" or request("FMEANo") ="" or request("ProjectName") ="" then
              Response.Write "<script language='JavaScript' type='text/javascript'>alert('δ�����������');</script>"
              Response.write "δ������룬��ȷ��.[<a href='javascript:history.back()'>����</a>]" 
              response.end
           end if   

           Pj_Key = trim(request("Serial")) & "-" &trim(request("FMEANo")) & "-" & trim(request("FMEARevision"))
           sql= "Insert into t_Fmea_PJ"
           sql= sql & "(PjKey,Serial,FMEANo,FMEARevision,ProjectName,ItemName,Model,Author,"
           sql= sql & "BurdenDept,Onduty,CreateDate,Pjstatus) values ('" 
           sql= sql & Pj_Key &"','" 
           sql= sql & request("Serial") &"','" 
           sql= sql & request("FMEANo") &"','" 
           sql= sql & request("FMEARevision") &"','" 
           sql= sql & request("ProjectName") &"','" 
           sql= sql & request("ItemName") &"','" 
           sql= sql & request("Model") &"','" 
           sql= sql & request("Author") &"','"  
           sql= sql & request("BurdenDept") &"','"
           sql= sql & request("Onduty") &"','"
           sql= sql & request("CreateDate") &"','Y')" 
           call openDB()
           conn.execute(sql)
           call closeDB()
           Response.Redirect "Pj_Browse.asp"				       
End Select

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������Ŀ��¼</title>
</head>

<body   marginheight="0" topmargin="10" bottommargin="0" leftmargin="0" bgcolor="#f9f9f9">
<form action="Pj_Add.asp" method="post">
<table width="80%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center">
<tr bgcolor="#6c9c6c">
<td colspan ="8" height ="30" align ="center">
<font size ="3" color ="ffffff">������Ŀ��¼</font></td></tr>
<tr bgcolor="#FCfCFC">
<td width = "80">����ϵ��</td>
<td><input name="Serial" class="notline" size="25">
<td width = "80">FMEA���</td>
<td><input name="FMEANo" class="notline" size="25">
<td  width = "80">�汾</td>
<td  colspan="3"><input name="FMEARevision" class="notline" size="25" value ="0A"></td>
</tr>
<tr bgcolor="#FCfCFC">
<td width = "80">��Ŀ</td>
<td><input name="ProjectName" class="notline" size="25"></td>
<td width = "80">Unit/��������</td>
<td><input name="ItemName" class="notline" size="25"></td>
<td width = "80">Model</td>
<td><input name="Model" class="notline" size="25"></td>
<td width = "80">������</td>
<td><input name="Author" class="notline" size="25"></td>
</tr>
<tr bgcolor="#FCfCFC">
<td  width = "80">��������</td>
<td><input name="BurdenDept" class="notline" size="25" value ="��������QC������QC3��"></td>
<td  width = "80">��Ҫ������</td>
<td><input name="Onduty" class="notline" size="25"></td>
<td  width = "80">FMEA��ʼ����</td>
<td><input name="CreateDate" class="notline" size="25" value ="<% = now() %>"></td>
<td  width = "80"></td>
<td></td>
</tr>
<tr bgcolor="#FCfCFC" height ="30">
<td colspan = "8" align ="center" >
<input name="OK" value="��¼" type="submit" class="smallInput">
</td></tr>
</table>
</form>
</body>
</html>

