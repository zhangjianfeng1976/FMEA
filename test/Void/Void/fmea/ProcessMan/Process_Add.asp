<!--#include file="../includes/db.asp"-->
<%
if  Request("PjID") = ""  then
    Response.Write "<script language='JavaScript' type='text/javascript'>alert('û��ָ��FMEA No.');</script>"    
    Response.Write "<a href ='..\PjManage\Pj_Browse.asp'>����û��ָ��FMEA No.���ԣ���ָ��FMEA No.��¼���������</a>"
    Response.End
end if
   
Select Case Request("OK")  
       Case "��Ŀ����"

       if Request("PjID") = ""  then
           Response.Write "<a href ='..\PjManage\Pj_Browse.asp'>������˼�����ݳ��ֶ�ʧ.����ָ��FMEA No.��¼��</a>"
           Response.End
       end if
 
       Sev = int(request("Sev"))
       Occ = int(request("Occ"))
       Det = int(request("Det"))
       RPN = Sev * Occ * Det
       sql= "Insert into t_Fmea_Item"
       sql= sql & "(PjID,SubPjName,PageNo,WorkQueue,PartsNo,PartsName,PartsType,WorkCata,WorkContent,"
       sql= sql & "NGMode,NGEffect,ProDesign,Sev,Occ,Det,RPN,ItmStatus,InUser,InTime) values ('" 
       sql= sql & request("PjID")&"','" 
       sql= sql & request("SubPjName")&"','" 
       sql= sql & request("PageNo")&"','" 
       sql= sql & request("WorkQueue")&"','"
       sql= sql & request("PartsNo")&"','" 
       sql= sql & request("PartsName")&"','" 
       sql= sql & request("PartsType")&"','" 
       sql= sql & request("WorkCata")&"','" 
       sql= sql & request("WorkContent")&"','"  
       sql= sql & request("NGMode")&"','"
       sql= sql & request("NGEffect")&"','"
       sql= sql & request("ProDesign")&"','"
       sql= sql & request("Sev")&"','"
       sql= sql & request("Occ")&"','"
       sql= sql & request("Det")&"','"
       sql= sql & RPN &"','Y','"
       sql= sql & request.cookies("fmea.truename") &"','"
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
<title>��Ŀ����</title>
</head>

<body   marginheight="0" topmargin="10" bottommargin="0" leftmargin="0" bgcolor="#f9f9f9">
<form action="Process_Add.asp" method="post">
<table width="66%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center">
<tr bgcolor="#6c9c6c">
<td colspan ="2" height ="30" align ="center">
<font size ="3" color ="ffffff"><% = Request("PjKey") %>������������</font>
<input name="PjID" type ="hidden" value ="<% = Request("PjID") %>">
<input name="PjKey" type ="hidden" value ="<% = Request("PjKey") %>">
</td></tr>
<tr bgcolor="#FCfCFC">
<td width = "150">��������</td>
<td><input name="SubPjName" class="notline" size="40">
</td></tr>
<tr bgcolor="#FCfCFC">
<td width = "150">��ҵ��׼��ҳ��</td>
<td><input name="PageNo" class="notline" size="25">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>��ҵ˳��</td>
<td><input name="WorkQueue" class="notline" size="5">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>��Ʒ���</td>
<td><input name="PartsNo" class="notline" size="25"></td></tr>
<tr bgcolor="#FCfCFC">
<td>��Ʒ����</td>
<td><input name="PartsName" class="notline" size="70"></td></tr>
<td bgcolor="#FCfCFC">��Ʒ����</td>
<td bgcolor="#FCfCFC">
<div>
<input name="PartsType" type ="Radio" value ="���"> ���
<input name="PartsType" type ="Radio" value ="�ܽ�"> �ܽ�
<input name="PartsType" type ="Radio" value ="����"> ����
<input name="PartsType" type ="Radio" value ="����"> ����
<input name="PartsType" type ="Radio" value ="�й��">�й��
<input name="PartsType" type ="Radio" value ="��������">��������
</div><div>
<input name="PartsType" type ="Radio" value ="����"> ����
<input name="PartsType" type ="Radio" value ="����"> ����
<input name="PartsType" type ="Radio" value ="���"> ���
<input name="PartsType" type ="Radio" value ="��Ӧ��"> ��Ӧ��
<input name="PartsType" type ="Radio" value ="�����"> �����
</div><div>
<input name="PartsType" type ="Radio" value ="���"> ���
<input name="PartsType" type ="Radio" value ="����"> ����
<input name="PartsType" type ="Radio" value ="��ǩ"> ��ǩ
<input name="PartsType" type ="Radio" value ="��˿"> ��˿
<input name="PartsType" type ="Radio" value ="��"> ��
<input name="PartsType" type ="Radio" value ="�Ĳ�"> �Ĳ�
</div>
</td></tr>
<tr bgcolor="#FCfCFC">
<td>��ҵ����</td>
<td>
<select name="WorkCata" class="notline">
 <option value="����˿">����˿</option>
 <option value="װ��">װ��</option>
 <option value="����">����</option>
 <option value="����">����</option>
 <option value="Ϳ��">Ϳ��</option>
 <option value="��ɨ">��ɨ</option>
 <option value="�趨">�趨</option>
 <option value="�ߴ���">�ߴ���</option>
</select>
</td></tr>
<tr bgcolor="#FCfCFC">
<td>��ҵ����</td>
<td><input name="WorkContent" class="notline" size="100"></td>
</tr>
<tr bgcolor="#FCfCFC">
<td>Ǳ��ȱ��ģʽ</td>
<td><input name="NGMode" class="notline" size="100"></td></tr>
<tr bgcolor="#FCfCFC">
<td>Ǳ��ȱ�ݺ��</td>
<td><input name="NGEffect" class="notline" size="100"></td></tr>
<tr bgcolor="#FCfCFC">
<td>���й�����Ƽ����̿���</td>
<td><input name="ProDesign" class="notline" size="100"></td></tr>
<tr bgcolor="#FCfCFC">
<td>���ض�Sev</td>
<td>
<input name="Sev" type ="Radio" value ="1"> 1
<input name="Sev" type ="Radio" value ="2"> 2
<input name="Sev" type ="Radio" value ="3"> 3
<input name="Sev" type ="Radio" value ="4"> 4
<input name="Sev" type ="Radio" value ="5"> 5
<input name="Sev" type ="Radio" value ="6"> 6
<input name="Sev" type ="Radio" value ="7"> 7
<input name="Sev" type ="Radio" value ="8"> 8
<input name="Sev" type ="Radio" value ="9"> 9
<input name="Sev" type ="Radio" value ="10"> 10
</td></tr>
<tr bgcolor="#FCfCFC">
<td>Ƶ��Occ</td>
<td>
<input name="Occ" type ="Radio" value ="1"> 1
<input name="Occ" type ="Radio" value ="2"> 2
<input name="Occ" type ="Radio" value ="3"> 3
<input name="Occ" type ="Radio" value ="4"> 4
<input name="Occ" type ="Radio" value ="5"> 5
<input name="Occ" type ="Radio" value ="6"> 6
<input name="Occ" type ="Radio" value ="7"> 7
<input name="Occ" type ="Radio" value ="8"> 8
<input name="Occ" type ="Radio" value ="9"> 9
<input name="Occ" type ="Radio" value ="10"> 10
</td></tr>
<tr bgcolor="#FCfCFC">
<td>����̽���Det</td>
<td>
<input name="Det" type ="Radio" value ="1"> 1
<input name="Det" type ="Radio" value ="2"> 2
<input name="Det" type ="Radio" value ="3"> 3
<input name="Det" type ="Radio" value ="4"> 4
<input name="Det" type ="Radio" value ="5"> 5
<input name="Det" type ="Radio" value ="6"> 6
<input name="Det" type ="Radio" value ="7"> 7
<input name="Det" type ="Radio" value ="8"> 8
<input name="Det" type ="Radio" value ="9"> 9
<input name="Det" type ="Radio" value ="10"> 10
</td>
</tr>
<tr bgcolor="#FCfCFC" height ="30">
<td colspan = "2" align ="center" >
<input name="OK" value="��Ŀ����" type="submit" class="smallInput">
</td></tr>
</table>
</form>
</body>
</html>

