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
      Case "ȷ��" 
       if request("line") ="" or request("model") ="���֡�"or request("wogroup") ="" or request("workstation") ="" or request("workstation") ="δ" or request("workstation") ="������" or request("originworker")="" or len(request("originworker")) <> 6  then
        response.write "����δ��ɻ��д���[<a href='javascript:history.back()'>����</a>]" 
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
<title>��λ��Ա����</title>
</head>

<body>
<form action ="add.asp"  method="POST">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=11>
<b>��Ա������λ��������</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="25">����</td>
<td width="8%">Line</td>
<td width="8%">����</td>
<td width="8%">����</td>
<td width="4%">��λ</td>
<td width="8%">������</td>
<td width="8%">��Ч</td>
<td width="30%">����</td>
</tr>
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="25">
<input type="text" name="originworker" class="notline" size="17" value="<%=request("cardid")%>"></td>
<td width="8%">
<select size="1" name="line" class="smallInput"  onChange="javascript:change_select(this.options[this.selectedIndex].value)">
    <option value="0">Line��</option>
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
arrSelect[0] =new Array("���֡�",0);
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
      <option>���֡�</option>
</select>

<!--<input type="text" name="model"  class="notline" size="17">-->
</td>
<td width="4%">
<select name="wogroup" class="smallInput">
 <option value="ǰ">ǰ</option>
 <option value="����">����</option>
 <option value="����">����</option>
 <option value="��װ">��װ</option>
 <option value="�ӹ�">�ӹ�</option>
 <option value="���">���</option>
 <option value="DLP">DLP</option>
 <option value="DRUM">DRUM</option>
 <option value="FUSER">FUSER</option>
 <option value="ISU">ISU</option>
 <option value="��ת">��ת</option>
 <option value="IH">IH</option>
 <option value="LSU">LSU</option>
 <option value="OP">OP</option>
 <option value="CONVEYING">CONVEYING</option>
 <option value="RPS">RPS</option>
 <option value="MK">MK</option>
 <option value="����">����</option>
 <option value="ӡˢ">ӡˢ</option>
 <option value="DP">DP</option>
 <option value="PF">PF</option>
 <option value="PUNCH">PUNCH</option>
 <option value="����">����</option>
 <option value="����">����</option>
 <option value="��ͷ">��ͷ</option>
 <option value="�ջ�">�ջ�</option>
</select>
<!--<input type="text" name="wogroup"  class="notline" size="17">-->
</td>
<td width="8%"><input type="text" name="workstation"  class="notline" size="17"></td>
<td width="8%"><input type="text" name="changedate"  class="notline" size="17" value ="<%=date()%>"></td>
<td width="8%">Y</td>
<td width="8%">
<input type="submit" value="ȷ��" name="OK" class="smallinput"></td>
</tr>
</table>
</form>
</body>
</html>