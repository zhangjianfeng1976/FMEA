<!--#include file="../includes/db.asp"-->
<%
SqlStr_Q ="select ItmID from t_Fmea_Advice Where ItmID ="&Request("ItmID")
call openDB()
Set result=conn.execute(SqlStr_Q)
If not result.eof then 
   call closeDB()
   Response.write "该项目的改善措施己输入，请返回查看.[<a href='javascript:history.back()'>返回</a>]"
   Response.end 
End If  

Select Case Request("OK")  
       Case "措施输入"
           R_Sev = int(request("ResultS"))
           R_Occ = int(request("ResultO"))
           R_Det = int(request("ResultD"))
           R_RPN = R_Sev * R_Occ * R_Det
           sql= "Insert into t_Fmea_Advice"
           sql= sql & "(ItmID,AdvContent,Rspnser,FshDate,ActionContent,"
           sql= sql & "ResultS,ResultO,ResultD,ResultRPN,AdvStatus,InUser,InTime) values ('" 
           sql= sql & request("ItmID")&"','" 
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
<table width="60%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center">
<tr bgcolor="#6c9c6c">
<td colspan ="2" height ="30" align ="center">
<font size ="3" color ="ffffff"><% = request("PjKey") %><% = request("ItmID") %>改善措施内容输入</font>
<input name="ItmID" type ="hidden" value ="<% = request("ItmID") %>"></td></tr>
<tr bgcolor="#FCfCFC">
<td width = "150">建议采取的措施事项</td>
<td><input name="AdvContent" class="notline" size="100">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>责任者</td>
<td><input name="Rspnser" class="notline" size="25">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>目标完成日</td>
<td><input name="FshDate" class="notline" size="25">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>采取措施</td>
<td><input name="ActionContent" class="notline" size="100">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>措施后严重度Sev</td>
<td>
<input name="ResultS" type ="Radio" value ="1"> 1
<input name="ResultS" type ="Radio" value ="2"> 2
<input name="ResultS" type ="Radio" value ="3"> 3
<input name="ResultS" type ="Radio" value ="4"> 4
<input name="ResultS" type ="Radio" value ="5"> 5
<input name="ResultS" type ="Radio" value ="6"> 6
<input name="ResultS" type ="Radio" value ="7"> 7
<input name="ResultS" type ="Radio" value ="8"> 8
<input name="ResultS" type ="Radio" value ="9"> 9
<input name="ResultS" type ="Radio" value ="10"> 10
</td></tr>
<tr bgcolor="#FCfCFC">
<td>措施后频度Occ</td>
<td>
<input name="ResultO" type ="Radio" value ="1"> 1
<input name="ResultO" type ="Radio" value ="2"> 2
<input name="ResultO" type ="Radio" value ="3"> 3
<input name="ResultO" type ="Radio" value ="4"> 4
<input name="ResultO" type ="Radio" value ="5"> 5
<input name="ResultO" type ="Radio" value ="6"> 6
<input name="ResultO" type ="Radio" value ="7"> 7
<input name="ResultO" type ="Radio" value ="8"> 8
<input name="ResultO" type ="Radio" value ="9"> 9
<input name="ResultO" type ="Radio" value ="10">10
</td></tr>
<tr bgcolor="#FCfCFC">
<td>措施后不易探测度Det</td>
<td>
<input name="ResultD" type ="Radio" value ="1"> 1
<input name="ResultD" type ="Radio" value ="2"> 2
<input name="ResultD" type ="Radio" value ="3"> 3
<input name="ResultD" type ="Radio" value ="4"> 4
<input name="ResultD" type ="Radio" value ="5"> 5
<input name="ResultD" type ="Radio" value ="6"> 6
<input name="ResultD" type ="Radio" value ="7"> 7
<input name="ResultD" type ="Radio" value ="8"> 8
<input name="ResultD" type ="Radio" value ="9"> 9
<input name="ResultD" type ="Radio" value ="10"> 10
</td>
</tr>
<tr bgcolor="#FCfCFC" height ="30">
<td colspan = "2" align ="center" >
<input name="OK" value="措施输入" type="submit" class="smallInput">
</td></tr>
</table>
</form>
</body>
</html>

