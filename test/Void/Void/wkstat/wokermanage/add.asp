<!--#include file="../includes/db.asp"-->
<%
if request.cookies("leader")="" then
response.redirect "../default.asp"
end if

sql1="select distinct Line from v_product_master"
sql2="SELECT * FROM v_product_master"

call openDB()
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")

rs1.open sql1,conn,1,1
rs2.open sql2,conn,1,1


Select Case Request("OK") 
      Case "确定" 
       if request("line") ="" or request("model") ="机种…"or request("wogroup") ="" or request("workstation") ="" or request("workstation") ="未" or request("workstation") ="工程外" or request("originworker")="" or len(request("originworker")) <> 6  then
        response.write "资料未完成或有错误[<a href='javascript:history.back()'>返回</a>]" 
        response.end
       end if 
       dim  statuspn
       statuspn=year(now())&right("0"&month(now()),2)&right("0"&day(now()),2)
       sql= "Insert into wt_lineconfig(line,model,wogroup,workstation,originworker,changedate,wkstatus,statuspno,leader,truename,intime) values ('" _
              &request("line")&"','" _
              &request("model")&"','" _
              &request("wogroup")&"','" _
              &request("workstation")&"','" _
              &request("originworker")&"','" _
              &request("changedate")&"','Y','" _
              &statuspn&"','"&Request.Cookies("leader")&"','"&Request.ServerVariables("REMOTE_HOST")&"','"&NOW()&"')"
      call openDB()
	 conn.execute(sql)
      call closeDB()
      Response.Redirect "../lineconfig/lcbrowse.asp?Cat=0&Sc="&Request("originworker")
End Select



%>
<html>
<head>
<title>工位人员管理</title>
</head>

<body>
<form action ="add.asp"  method="POST">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=11>
<b>人员担当工位数据输入</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="25">工号</td>
<td width="8%">Line</td>
<td width="8%">机种</td>
<td width="8%">工程</td>
<td width="4%">工位</td>
<td width="8%">担当日</td>
<td width="8%">有效</td>
<td width="30%">安排</td>
</tr>
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="25">
<input type="text" name="originworker" class="notline" size="17" value="<%=request("cardid")%>"></td>
<td width="8%">
<select size="1" name="line" class="smallInput"  onChange="javascript:change_select(this.options[this.selectedIndex].value)">
    <option value="0">Line…</option>
<%
If rs1.eof then
End if

n=1 
While Not rs1.eof 
  response.write"<option value="&rs1("Line")&">"&rs1("Line")&"</option>"
  n=n+1 
  rs1.movenext  
Wend 
%>
  </select>

<!--<input type="text" name="line"  class="notline" size="17">-->
</td>
<td width="8%">
<select name="model" class="smallInput" >
<script language = "JavaScript"> 
arrSelect=new Array();
arrSelect[0] =new Array("机种…",0);
<%
If rs2.eof then
End if

i=1 
While Not rs2.eof 
  response.write "arrSelect["&i&"] =new Array('"&trim(rs2("Model"))&"','"&trim(rs2("Line"))&"');"
  i=i+1 
  rs2.movenext  
Wend 
%>

function change_select(selvalue) 
{ 
document.all.model.length=0; 
var i;
for(i=0;i<arrSelect.length;i++) 
{
if(arrSelect[i][1]==selvalue) 
{
var newOption=new Option(arrSelect[i][0]);
document.all.model.add(newOption);
}
}
}
</script>
      <option>机种…</option>
</select>

<!--<input type="text" name="model"  class="notline" size="17">-->
</td>
<td width="4%">
<select name="wogroup" class="smallInput">
 <option value="前">前</option>
 <option value="调整">调整</option>
 <option value="仕上">仕上</option>
 <option value="包装">包装</option>
 <option value="加工">加工</option>
 <option value="检查">检查</option>
 <option value="DLP">DLP</option>
 <option value="DRUM">DRUM</option>
 <option value="FUSER">FUSER</option>
 <option value="ISU">ISU</option>
 <option value="中转">中转</option>
 <option value="IH">IH</option>
 <option value="LSU">LSU</option>
 <option value="OP">OP</option>
 <option value="CONVEYING">CONVEYING</option>
 <option value="RPS">RPS</option>
 <option value="MK">MK</option>
 <option value="画像">画像</option>
 <option value="印刷">印刷</option>
 <option value="DP">DP</option>
 <option value="PF">PF</option>
 <option value="PUNCH">PUNCH</option>
 <option value="配膳">配膳</option>
 <option value="车手">车手</option>
 <option value="码头">码头</option>
 <option value="日货">日货</option>
</select>
<!--<input type="text" name="wogroup"  class="notline" size="17">-->
</td>
<td width="8%"><input type="text" name="workstation"  class="notline" size="17"></td>
<td width="8%"><input type="text" name="changedate"  class="notline" size="17" value ="<%=date()%>"></td>
<td width="8%">Y</td>
<td width="8%">
<input type="submit" value="确定" name="OK" class="smallinput"></td>
</tr>
</table>
</form>
</body>
</html>