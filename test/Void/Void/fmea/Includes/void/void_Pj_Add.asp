<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK")  
       Case "登录"
           if request("Serial") ="" or request("FMEANo") ="" or request("ProjectName") ="" then
              Response.Write "<script language='JavaScript' type='text/javascript'>alert('未完成输入内容');</script>"
              Response.write "未完成输入，请确认.[<a href='javascript:history.back()'>返回</a>]" 
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
<title>机种项目登录</title>
</head>

<body   marginheight="0" topmargin="10" bottommargin="0" leftmargin="0" bgcolor="#f9f9f9">
<form action="Pj_Add.asp" method="post">
<table width="80%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center">
<tr bgcolor="#6c9c6c">
<td colspan ="8" height ="30" align ="center">
<font size ="3" color ="ffffff">机种项目登录</font></td></tr>
<tr bgcolor="#FCfCFC">
<td width = "80">机种系列</td>
<td><input name="Serial" class="notline" size="25">
<td width = "80">FMEA编号</td>
<td><input name="FMEANo" class="notline" size="25">
<td  width = "80">版本</td>
<td  colspan="3"><input name="FMEARevision" class="notline" size="25" value ="0A"></td>
</tr>
<tr bgcolor="#FCfCFC">
<td width = "80">项目</td>
<td><input name="ProjectName" class="notline" size="25"></td>
<td width = "80">Unit/本体名称</td>
<td><input name="ItemName" class="notline" size="25"></td>
<td width = "80">Model</td>
<td><input name="Model" class="notline" size="25"></td>
<td width = "80">作成者</td>
<td><input name="Author" class="notline" size="25"></td>
</tr>
<tr bgcolor="#FCfCFC">
<td  width = "80">过程责任</td>
<td><input name="BurdenDept" class="notline" size="25" value ="机器工程QC部工程QC3科"></td>
<td  width = "80">主要责任者</td>
<td><input name="Onduty" class="notline" size="25"></td>
<td  width = "80">FMEA初始日期</td>
<td><input name="CreateDate" class="notline" size="25" value ="<% = now() %>"></td>
<td  width = "80"></td>
<td></td>
</tr>
<tr bgcolor="#FCfCFC" height ="30">
<td colspan = "8" align ="center" >
<input name="OK" value="登录" type="submit" class="smallInput">
</td></tr>
</table>
</form>
</body>
</html>

