<!--#include file="../includes/db.asp"-->
<%
if  Request.cookies("fmea.PjID") = ""  then
    Response.Write "<script language='JavaScript' type='text/javascript'>alert('û��ָ��FMEA No.');</script>"    
    Response.Write "<a href ='..\PjManage\Pj_Browse.asp'>����û��ָ��FMEA No.���ԣ���ָ��FMEA No.��¼���������</a>"
    Response.End
end if
   
Select Case Request("OK")  
       Case "�ص��б�"
           Response.redirect "Process_Browse.asp"

       Case "��Ŀ����"

       if Request.cookies("fmea.PjID") = ""  then
           Response.Write "<a href ='..\PjManage\Pj_Browse.asp'>������˼�����ݳ��ֶ�ʧ.����ָ��FMEA No.��¼��</a>"
           Response.End
       end if
       
       sql_ck = "select * from t_fmea_Item where PjID = '"&Request.cookies("fmea.PjID")&"'"
       sql_ck = sql_ck&" and  SubPjName ='"&request("SubPjName")&"'"
       sql_ck = sql_ck&" and PageNo='"&request("PageNo")&"'"
       sql_ck = sql_ck&" and WorkQueue='"&request("WorkQueue")&"'"
       
       call openDB()
       rs.open sql_ck,conn,1,1

       If rs.eof then
         Sev = int(request("Sev"))
         Occ = int(request("Occ"))
         Det = int(request("Det"))
         RPN = round((Sev * Occ * Det)^(1/3),1)
     
         ProDesign =request("Pro1")&request("Pro2")&request("Pro3")&request("Pro4")&request("ProDesign")

         sql= "Insert into t_Fmea_Item"
         sql= sql & "(PjID,Pjkey,SubPjName,PageNo,WorkQueue,WorkTheme,PartsNo,PartsName,PartsType,WorkCata,WorkContent,"
         sql= sql & "NGMode,NGEffect,ProDesign,Sev,Occ,Det,RPN,ItmStatus,InUser,InTime) values ('" 
         sql= sql & Request.cookies("fmea.PjID")&"','" 
         sql= sql & Request.cookies("fmea.Pjkey")&"','" 
         sql= sql & request("SubPjName")&"','" 
         sql= sql & request("PageNo")&"','" 
         sql= sql & request("WorkQueue")&"','"
		 sql= sql & request("WorkTheme")&"','" 
         sql= sql & request("PartsNo")&"','" 
         sql= sql & request("PartsName")&"','" 
         sql= sql & request("PartsType")&"','" 
         sql= sql & request("WorkCata")&"','" 
         sql= sql & request("WorkContent")&"','"  
         sql= sql & request("NGMode")&"','"
         sql= sql & request("NGEffect")&"','"
         sql= sql & ProDesign&"','"
         sql= sql & request("Sev")&"','"
         sql= sql & request("Occ")&"','"
         sql= sql & request("Det")&"','"
         sql= sql & RPN &"','Y','"
         sql= sql & request.cookies("fmea.truename") &"','"
         sql= sql & Now() &"')" 
    
         conn.execute(sql)
         call closeDB()
        
       Else
         response.write "<script language='JavaScript'>alert('����ҳ�뼰��ҵ˳���ظ�!');</script>"
         Response.Write "<script language='javascript' type='text/javascript'>history.go(-1)</script>"
         response.end
       End if 
				       
End Select
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript" type="text/javascript">
<!--
function jump() {
 if(event.keyCode==13)event.keyCode=9
}

function showCustomer(str){
var xmlhttp;
if (str=="")
 {
 document.getElementById("txtHint").innerHTML="";
 return;
 }
if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
 xmlhttp=new XMLHttpRequest();
 }
else {// code for IE6, IE5
 xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
 }
xmlhttp.onreadystatechange=function() {
 if (xmlhttp.readyState==4 && xmlhttp.status==200) {
  document.getElementById("txtHint").innerHTML=xmlhttp.responseText;
    }
 }
xmlhttp.open("GET","getcustomer.html?q="+str,true);
xmlhttp.send();
}

-->
</script>
<title>��Ŀ����</title>
</head>

<body marginheight="0" topmargin="10" bottommargin="0" leftmargin="0" bgcolor="#f9f9f9">
<form action="Process_Add.asp" method="post">
<table width="750"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center">
<tr bgcolor="#6c9c6c">
<td colspan ="2" height ="30" align ="center">
<font size ="3" color ="ffffff"><% = Request.cookies("fmea.PjKey") %>������������</font>
</td></tr>
<tr bgcolor="#FCFCFC">
<td width = "150">��������</td>
<td><input name="SubPjName" class="notline" size="40" value ="<% = Request("SubPjName") %>" tabindex ="1" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCFCFC">
<td width = "150">��ҵ��׼��ҳ��</td>
<td><input name="PageNo" class="notline" size="25" value ="<% = Request("PageNo") %>" tabindex ="2" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCFCFC">
<td>��ҵ˳��</td>
<td><input name="WorkQueue" class="notline" size="8" value ="<% = Request("WQ") %>"  tabindex ="3" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCFCFC">
<td>��ҵ����</td>
<td>
<% 
  response.write"<select name ='WorkTheme'>"
  For i = 0 to UBound(ArWkThm) 
     response.write"<option value ='"&ArWkThm(i)&"'>"&ArWkThm(i)&"</option>"
  Next
  response.write"</select>"
%>  
<!-- <input name="WorkTheme" class="notline" size="50"  tabindex ="4" onKeyDown="jump()"> -->

</td></tr>
<tr bgcolor="#FCFCFC">
<td>��Ʒ���</td>
<td><input name="PartsNo" class="notline" size="50"  tabindex ="5" onKeyDown="jump()"></td></tr>
<tr bgcolor="#FCFCFC">
<td>��Ʒ����</td>
<td><input name="PartsName" class="notline" size="70"  tabindex ="6" onKeyDown="jump()"></td></tr>
<td bgcolor="#FCFCFC">��Ʒ����</td>
<td bgcolor="#FCFCFC">
<div>
<input name="PartsType" type ="Radio" value ="��ܼ�"  tabindex ="7" onKeyDown="jump()"> ��ܼ�
<input name="PartsType" type ="Radio" value ="������"> ������
<input name="PartsType" type ="Radio" value ="��ѧ��"> ��ѧ��
<input name="PartsType" type ="Radio" value ="������"> ������
<input name="PartsType" type ="Radio" value ="������"> ������
<input name="PartsType" type ="Radio" value ="��ͨ��"> ��ͨ��
<input name="PartsType" type ="Radio" value ="��������"> ��������
</div><div>
<input name="PartsType" type ="Radio" value ="����"> ����
<input name="PartsType" type ="Radio" value ="���"> ���
<input name="PartsType" type ="Radio" value ="����"> ����
<input name="PartsType" type ="Radio" value ="��Ƭ"> ��Ƭ
<input name="PartsType" type ="Radio" value ="����"> ����
</div><div>
<input name="PartsType" type ="Radio" value ="��ǩ"> ��ǩ
<input name="PartsType" type ="Radio" value ="��˿"> ��˿
<input name="PartsType" type ="Radio" value ="ͬ��Ʒ"> ͬ��Ʒ
<input name="PartsType" type ="Radio" value ="�󻬲�"> �󻬲�
<input name="PartsType" type ="Radio" value ="�Ĳ�"> �Ĳ�
<input name="PartsType" type ="Radio" value ="����"> ����
<input name="PartsType" type ="Radio" value ="-"> -
</div>
</td></tr>
<tr bgcolor="#FCFCFC">
<td>��ҵ����</td>
<td>
<select name="WorkCata" class="notline"  tabindex ="8" onKeyDown="jump()">
 <option value="����˿">����˿</option>
 <option value="װ��">װ��</option>
 <option value="����">����</option>
 <option value="����">����</option>
 <option value="�ߴ���">�ߴ���</option> 
 <option value="Ϳ��">Ϳ��</option>
 <option value="��ɨ">��ɨ</option>
 <option value="�趨">�趨</option>
 <option value="���">���</option>
 <option value="ͬ��">ͬ��</option>
 <option value="-">-</option>
</select>
</td></tr>
<tr bgcolor="#FCFCFC">
<td>��ҵ����</td>
<td><input name="WorkContent" class="notline" size="100"  tabindex ="9" onKeyDown="jump()"></td>
</tr>
<tr bgcolor="#FCFCFC">
<td>Ǳ��ȱ��ģʽ</td>
<td><input name="NGMode" class="notline" size="100"  tabindex ="10" onKeyDown="jump()"></td></tr>
<tr bgcolor="#FCFCFC">
<td>Ǳ��ȱ�ݺ��</td>
<td><input name="NGEffect" class="notline" size="100"  tabindex ="11" onKeyDown="jump()"></td></tr>
<tr bgcolor="#FCFCFC">
<td>���й�����Ƽ����̿���</td>
<td>
<input name="Pro1" type ="checkbox" value ="���̹���">���̹���
<input name="Pro2" type ="checkbox" value ="˳�ι���">˳�ι���
<input name="Pro3" type ="checkbox" value ="��������">��������
<input name="Pro4" type ="checkbox" value ="��������">��������
&nbsp;&nbsp;������<input name="ProDesign" class="notline" size="30" value ="Ŀ�Ӽ��" tabindex ="12" onKeyDown="jump()">

</td></tr>
<tr bgcolor="#FCFCFC">
<td>���ض�Sev</td>
<td>
<input name="Sev" type ="Radio" value ="1"  tabindex ="13" onKeyDown="jump()"> 1
<input name="Sev" type ="Radio" value ="2"> 2
<input name="Sev" type ="Radio" value ="3"> 3
<input name="Sev" type ="Radio" value ="4"> 4

</td></tr>
<tr bgcolor="#FCFCFC">
<td>Ƶ��Occ</td>
<td>
<input name="Occ" type ="Radio" value ="1"  tabindex ="14" onKeyDown="jump()"> 1
<input name="Occ" type ="Radio" value ="2"> 2
<input name="Occ" type ="Radio" value ="3"> 3
<input name="Occ" type ="Radio" value ="4"> 4

</td></tr>
<tr bgcolor="#FCFCFC">
<td>����̽���Det</td>
<td>
<input name="Det" type ="Radio" value ="1"  tabindex ="15" onKeyDown="jump()"> 1
<input name="Det" type ="Radio" value ="2"> 2
<input name="Det" type ="Radio" value ="3"> 3
<input name="Det" type ="Radio" value ="4"> 4

</td>
</tr>
<tr bgcolor="#FCFCFC" height ="30">
<td colspan = "2" align ="center" >
<input name="OK" value="��Ŀ����" type="submit" class="smallInput"  tabindex ="16">
<input name="OK" value="�ص��б�" type="submit" class="smallInput"  tabindex ="17">
</td></tr>
</table>
</form>
</body>
</html>

