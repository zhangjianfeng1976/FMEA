<!--#include file="../includes/db.asp"-->
<%
if (request.cookies("fmea.serial") ="" or request.cookies("fmea.serial") <> request("serial")) and request("serial")<> "" then
response.cookies("fmea.serial") = request("serial")
end if

Select Case Request("OK") 
    Case "����Ŀ��¼"
         Response.Redirect "Pj_Add.asp"

    Case "����" 	                
	     Response.Redirect "Pj_Browse.asp?Cat=0&Sc="&Request("Search")
       
    Case "ȫ��ɾ��"
         sql_ad = "Insert into t_Fmea_Advice_history "
         sql_ad = sql_ad & "(AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
         sql_ad = sql_ad & "ActionContent,ResultS,ResultO,ResultD,"
         sql_ad = sql_ad & "ResultRPN,AdvStatus,InUser,InTime)"
         sql_ad = sql_ad & " Select AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
         sql_ad = sql_ad & "ActionContent,ResultS,ResultO,ResultD,"
         sql_ad = sql_ad & "ResultRPN,'N',InUser,InTime"
         sql_ad = sql_ad & " from t_Fmea_Advice Where PjID = '"& request("ID") & "'"

         sql_ih = "Insert into t_Fmea_Item_history "
         sql_ih = sql_ih & "(ItmID,PjID,Pjkey,SubPjName,PageNo,WorkQueue,PartsNo,PartsName,PartsType,WorkCata,"
         sql_ih = sql_ih & "WorkContent,NGMode,NGEffect,ProDesign,Sev,Occ,"
         sql_ih = sql_ih & "Det,RPN,ItmStatus,InUser,InTime)"
         sql_ih = sql_ih & " Select ItmID, PjID ,Pjkey,SubPjName,PageNo,WorkQueue,PartsNo,"
         sql_ih = sql_ih & "PartsName,PartsType,WorkCata,WorkContent,NGMode,NGEffect,"
         sql_ih = sql_ih & "ProDesign,Sev,Occ,Det,RPN,'N',InUser,InTime"
         sql_ih = sql_ih & " from t_Fmea_Item Where PjID = '"& request("ID") & "'"
  
         sql_bl = "Insert into t_Fmea_PJ_history "
         sql_bl = sql_bl & "(PjID,Pjkey,Serial,FMEANo,FMEARevision,ProjectName,ItemName,Model,"
         sql_bl = sql_bl & "Author,BurdenDept,Onduty,CreateDate,PjStatus,ModifyDate,"
         sql_bl = sql_bl & "ApplyPerson,ApplyDate,AuditPerson,AuditDate,"
         sql_bl = sql_bl & "ApprovePerson,ApproveDate,PublishPerson,PublishDate,"
         sql_bl = sql_bl & "InUser,InTime) "
         sql_bl = sql_bl & "Select PjID,Pjkey,Serial,FMEANo,FMEARevision,ProjectName,"
         sql_bl = sql_bl & "ItemName,Model,Author,BurdenDept,Onduty,CreateDate,'DL',"
         sql_bl = sql_bl & "'" &Now()& "',ApplyPerson,ApplyDate,AuditPerson,AuditDate,"
         sql_bl = sql_bl & "ApprovePerson,ApproveDate,PublishPerson,PublishDate,InUser,InTime"
         sql_bl = sql_bl & " From t_Fmea_PJ Where PjID = '"& request("ID")&"'"
         
         SqlStr_dadv = "Delete t_Fmea_Advice Where PjID = '"&request("ID")&"'"
        
         SqlStr_ditm = "Delete t_Fmea_Item Where PjID = '"&request("ID")&"'"
 
         SqlStr_d = "Delete t_Fmea_PJ Where PjID = '"&request("ID")&"'"
            
         call openDB() 
         conn.BeginTrans
         conn.execute(Sql_ad)  
         conn.execute(Sql_ih) 
         conn.execute(Sql_bl)
         conn.execute(SqlStr_dadv) 
         conn.execute(SqlStr_ditm) 
         conn.execute(SqlStr_d)
         conn.CommitTrans
         call closeDB()
         Response.Redirect "Pj_Browse.asp"

    Case "����ɾ��(����)"
         
         sql_ad = "Insert into t_Fmea_Advice_history "
         sql_ad = sql_ad & "(AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
         sql_ad = sql_ad & "ActionContent,ResultS,ResultO,ResultD,"
         sql_ad = sql_ad & "ResultRPN,AdvStatus,InUser,InTime)"
         sql_ad = sql_ad & " Select AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
         sql_ad = sql_ad & "ActionContent,ResultS,ResultO,ResultD,"
         sql_ad = sql_ad & "ResultRPN,'N',InUser,InTime"
         sql_ad = sql_ad & " from t_Fmea_Advice Where PjID = '"& request("ID") & "'"
             
         sql_ih = "Insert into t_Fmea_Item_history "
         sql_ih = sql_ih & "(ItmID,PjID,Pjkey,SubPjName,PageNo,WorkQueue,PartsNo,PartsName,PartsType,WorkCata,"
         sql_ih = sql_ih & "WorkContent,NGMode,NGEffect,ProDesign,Sev,Occ,"
         sql_ih = sql_ih & "Det,RPN,ItmStatus,InUser,InTime)"
         sql_ih = sql_ih & " Select ItmID, PjID ,Pjkey,SubPjName,PageNo,WorkQueue,PartsNo,"
         sql_ih = sql_ih & "PartsName,PartsType,WorkCata,WorkContent,NGMode,NGEffect,"
         sql_ih = sql_ih & "ProDesign,Sev,Occ,Det,RPN,'N',InUser,InTime"
         sql_ih = sql_ih & " from t_Fmea_Item Where PjID = '"& request("ID") & "'"

         sql_bl = "Insert into t_Fmea_PJ_history "
         sql_bl = sql_bl & "(PjID,Pjkey,Serial,FMEANo,FMEARevision,ProjectName,ItemName,Model,"
         sql_bl = sql_bl & "Author,BurdenDept,Onduty,CreateDate,PjStatus,ModifyDate,"
         sql_bl = sql_bl & "ApplyPerson,ApplyDate,AuditPerson,AuditDate,"
         sql_bl = sql_bl & "ApprovePerson,ApproveDate,PublishPerson,PublishDate,"
         sql_bl = sql_bl & "InUser,InTime) "
         sql_bl = sql_bl & "Select PjID,Pjkey,Serial,FMEANo,FMEARevision,ProjectName,"
         sql_bl = sql_bl & "ItemName,Model,Author,BurdenDept,Onduty,CreateDate,'DL',"
         sql_bl = sql_bl & "'" &Now()& "',ApplyPerson,ApplyDate,AuditPerson,AuditDate,"
         sql_bl = sql_bl & "ApprovePerson,ApproveDate,PublishPerson,PublishDate,InUser,InTime"
         sql_bl = sql_bl & " From t_Fmea_PJ Where PjID = '"& request("ID")&"'"
         
         SqlStr_dadv = "Delete t_Fmea_Advice Where PjID = '"&request("ID")&"'"

         SqlStr_ditm = "Delete t_Fmea_Item Where PjID = '"&request("ID")&"'"
         
         call openDB() 
         conn.BeginTrans
         conn.execute(Sql_ad)  
         conn.execute(Sql_ih) 
         conn.execute(Sql_bl) 
         conn.execute(SqlStr_dadv) 
         conn.execute(SqlStr_ditm)
         conn.CommitTrans
         call closeDB()
         Response.Redirect "Pj_Browse.asp?act=verup&ID="&request("ID")		             

    Case "�޶�"
         Pj_Key = trim(request("nSerial")) & "-" &trim(request("FMEANo")) & "-" &trim(request("FMEARevision"))       
         SqlStr_c ="Update t_Fmea_PJ set "
         SqlStr_c = SqlStr_c & "PjKey = '"& Pj_Key& "',"
         SqlStr_c = SqlStr_c & "Serial = '"& request("nSerial") & "',"
         SqlStr_c = SqlStr_c & "FMEANo = '"& request("FMEANo") & "',"
         SqlStr_c = SqlStr_c & "FMEARevision = '"& request("FMEARevision") & "',"
         SqlStr_c = SqlStr_c & "ProjectName = '"& request("ProjectName") & "',"
         SqlStr_c = SqlStr_c & "ItemName = '"& request("ItemName") & "',"
         SqlStr_c = SqlStr_c & "Model = '"& request("Model") & "',"
         SqlStr_c = SqlStr_c & "Author = '"& request("Author") & "',"
         SqlStr_c = SqlStr_c & "BurdenDept = '"& request("BurdenDept") & "',"
         SqlStr_c = SqlStr_c & "Onduty = '"& request("Onduty") & "',"
         SqlStr_c = SqlStr_c & "CreateDate = '"& request("CreateDate") & "'"
         SqlStr_c = SqlStr_c & " Where PjID = '"& request("PjID") & "'"
         call openDB()
         conn.execute(SqlStr_c)
         call closeDB()
         Response.Redirect "Pj_Browse.asp"

    Case "����" 
        Pj_Key  = trim(request("nSerial")) & "-" &trim(request("FMEANo"))  & "-" &trim(request("FMEARevision"))
        Pjkey_C = request.cookies("fmea.Pjkey")
        nPjKey  = request("nPjkey")
        
        if request("PublishPerson") ="" then
           Response.Write "<script language='JavaScript' type='text/javascript'>alert('����δ���');</script>"
           Response.write "δ���ȫ���ĳ������̣����ܽ��������������ȷ��.[<a href='javascript:history.back()'>����</a>]" 
           Response.end
        end if
        
        if Pj_Key = nPjKey  then
            Response.Write "<script language='JavaScript' type='text/javascript'>alert('������');</script>"
            Response.write "û�н����������(�汾��δ�ģ�������ȷ��.[<a href='javascript:history.back()'>����</a>]" 
            Response.end
        end if
           
        sql_vu = "select Pjkey from t_Fmea_Advice_history where Pjkey = '"&nPjkey&"' order by PjID Desc"
            
        call openDB()
        set rs_vu = conn.execute(sql_vu)
        if  rs_vu.Eof then
             call closeDB()
             Response.write "δ�иĶ����ݣ���ȷ��.[<a href='javascript:history.back()'>����</a>]"
             Response.end           
            
        else
            call closeDB()
            
            '--������ʷ
            sql = "Insert into t_Fmea_PJ_history "
            sql = sql & "(PjID,Pjkey,Serial,FMEANo,FMEARevision,ProjectName,ItemName,Model,"
            sql = sql & "Author,BurdenDept,Onduty,CreateDate,PjStatus,ModifyDate,"
            sql = sql & "ApplyPerson,ApplyDate,AuditPerson,AuditDate,"
            sql = sql & "ApprovePerson,ApproveDate,PublishPerson,PublishDate,"
            sql = sql & "InUser,InTime)"
            sql = sql & "Select PjID,Pjkey,Serial,FMEANo,FMEARevision,ProjectName,"
            sql = sql & "ItemName,Model,Author,BurdenDept,Onduty,CreateDate,'N',"
            sql = sql & "'" &Now()& "',ApplyPerson,ApplyDate,AuditPerson,AuditDate,"
            sql = sql & "ApprovePerson,ApproveDate,PublishPerson,PublishDate,InUser,InTime"
            sql = sql & " from t_Fmea_PJ Where PjID = '"& request("PjID") & "'"

            '--�������
            
            SqlStr_c = "Update t_Fmea_PJ set "
            SqlStr_c = SqlStr_c & "PjKey = '"& Pj_Key& "',"
            SqlStr_c = SqlStr_c & "Serial = '"& request("nSerial") & "',"
            SqlStr_c = SqlStr_c & "FMEANo = '"& request("FMEANo") & "',"
            SqlStr_c = SqlStr_c & "FMEARevision = '"& request("FMEARevision") & "',"
            SqlStr_c = SqlStr_c & "ModifyDate = '"& Now() & "',"
            SqlStr_c = SqlStr_c & "ApplyPerson = Null,"
            SqlStr_c = SqlStr_c & "ApplyDate = Null,"
            SqlStr_c = SqlStr_c & "AuditPerson = Null,"
            SqlStr_c = SqlStr_c & "AuditDate = Null,"
            SqlStr_c = SqlStr_c & "ApprovePerson = Null,"
            SqlStr_c = SqlStr_c & "ApproveDate = Null,"
            SqlStr_c = SqlStr_c & "PublishPerson = Null,"
            SqlStr_c = SqlStr_c & "PublishDate = Null,"
            SqlStr_c = SqlStr_c & "InUser = '"&request.cookies("fmea.truename")&"',"
            SqlStr_c = SqlStr_c & "InTime = '" & now() &"'"
            SqlStr_c = SqlStr_c & " Where PjID = '"& request("PjID") & "'"
            
            '--���ĸ�ס�������
            SqlStr_Item = "Update t_Fmea_Item set "
            SqlStr_Item = SqlStr_Item & "PjKey = '"& Pj_Key& "'"
            SqlStr_Item = SqlStr_Item & " Where PjID = '"& request("PjID") & "'"
            
             '--��Ӧ��ʩ��ס�������
            SqlStr_Adv = "Update t_Fmea_Advice set "
            SqlStr_Adv = SqlStr_Adv & "PjKey = '"& Pj_Key& "'"
            SqlStr_Adv = SqlStr_Adv & " Where PjID = '"& request("PjID") & "'"

            call openDB()
            conn.BeginTrans
            conn.execute(sql)
            conn.execute(SqlStr_c) 
            conn.execute(SqlStr_Item) 
            conn.execute(SqlStr_Adv)           
            conn.CommitTrans
            call closeDB()
            
            Response.Redirect "Pj_Browse.asp"

            end if
End Select

pgs =2500
select case request("cat")
       case "" 
         ttl = request.cookies("fmea.serial")
         sql = "SELECT * FROM t_Fmea_PJ where Serial = '"&request.cookies("fmea.serial")&"' Order by PjKey" 
         pgs = 500
       case 0
         ttl = request.cookies("fmea.serial")
         sql="SELECT * FROM t_Fmea_PJ where (PjID like '%"&request("Sc")& _
           "%' Or Pjkey like '%"&request("Sc")& _
           "%' Or Serial like '%"&request("Sc")& _
           "%' Or FMEANo like '%"&request("Sc")& _
           "%' Or FMEARevision like '%"&request("Sc")& _
           "%' Or ProjectName like '%"&request("Sc")& _
           "%' Or ItemName like '%"&request("Sc")& _
           "%' Or Model like '%"&request("Sc")& _
           "%' Or Author like '%"&request("Sc")& _
           "%') and Serial = '"&request.cookies("fmea.serial")&"' Order by PjKey" 
      case 1
         ttl = "������"
         sql = "SELECT * FROM t_Fmea_PJ where  ApplyPerson IS NULL"
      case 2 
         ttl = "�����"
         sql = "SELECT * FROM t_Fmea_PJ where  ApplyPerson IS NOT NULL and AuditPerson IS NULL"
      case 3
         ttl = "������" 
         sql = "SELECT * FROM t_Fmea_PJ where  AuditPerson IS NOT NULL and ApprovePerson IS NULL"
      case 4 
         ttl = "���䲼" 
         sql = "SELECT * FROM t_Fmea_PJ where  ApprovePerson IS NOT NULL and PublishPerson IS NULL"
      case 5 
         ttl = "�䲼���" 
         sql = "SELECT * FROM t_Fmea_PJ where  PublishPerson IS NOT NULL"
end select

call openDB()
rs.open sql,conn,1,1
if not rs.eof then
rs.pagesize =pgs
pagecount=rs.PageCount 
page=int(request.QueryString ("page"))
if page<=0 then page=1
if request.QueryString("page")="" then page=1
rs.AbsolutePage=page 
end if
%>
<html>
<head>
<script language="javascript" type="text/javascript">
<!--
function jump() 
{
 if(event.keyCode==13)event.keyCode=9
}
-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ֹ���</title>
</head>

<body bgcolor ="#FCFCFC">
<form action="Pj_Browse.asp">
<div style ="float:right ;width��200px;hight:200px;">
<%
if request.cookies("fmea.power") >  "1" then 
   response.write "<input type='submit' value='����Ŀ��¼' name='OK' class='smallInput'>"
end if
%>
<input name="Search" type="text" class="smallInput" size="40">
<input type="submit" value="����" name="OK" class="smallInput">
&nbsp;&nbsp;
</div>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=13>
<% = ttl %><b>�����б�</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="5%">����״̬</td>
<td width="3%">����</td>
<td width="8%" height="20">FMEA��ż��汾</td>
<td width="10%">��Ŀ</td>
<td width="12%">����</td>
<td width="8%">����</td>
<td width="5%">������</td>
<td width="10%">��������</td>
<td width="8%">��Ҫ������</td>
<td width="8%">FMEA��ʼ����</td>
<%
if request.cookies("fmea.power") > "1" then
    response.write "<td width='10%'>����</td>"
end if
%>
</tr>

<% 
if request("act") = "del" then

    SqlStr = "SELECT * FROM t_Fmea_PJ where PjID ="&request("ID")
      
    Set rs_ed = Server.CreateObject("ADODB.Recordset")        
    rs_ed.open SqlStr,conn,1,1
       
    inuser = trim(rs_ed("InUser"))
    truename = trim(request.cookies("fmea.truename"))
    if inuser <> truename  then
	    response.write"<tr bgcolor='#FCFCFC'><td colspan ='11' height ='20'>"
        response.write"�㲻��ɾ�����˵ļ�¼!![<a href='javascript:history.back()'>����</a>]"
	    response.write"</td></tr>"
        response.End()
        Call closeDB()
    else
        response.write"<input name='ID' type='hidden' value ='"&rs_ed("PjID")&"'>"
        response.write"<tr bgcolor='#FCFCFC'><td colspan ='11'> ȷʵҪɾ��FMEA?&nbsp;&nbsp;&nbsp;&nbsp;"
        response.write"<input name='OK' value='ȫ��ɾ��' type='submit' class='smallInput'>&nbsp;&nbsp;&nbsp;&nbsp;"
        response.write"<input name='OK' value='����ɾ��(����)' type='submit' class='smallInput'>&nbsp;&nbsp;&nbsp;&nbsp;"
        response.write"[<a href='javascript:history.back()'>����</a>]</td></tr>"
    end if

elseif request("act") = "edit" then

    SqlStr_e = "SELECT * FROM t_Fmea_PJ where PjID ="&request("ID")
    Set rs_ed = Server.CreateObject("ADODB.Recordset")        
    rs_ed.open SqlStr_e,conn,1,1
    response.write"<tr bgcolor='#FCFCFC'><td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td>"
    response.write"<input name='PjID' type='hidden' value ='"&rs_ed("PjID")&"'>"
    response.write"<input name='Pjkey' type='hidden' value ='"&rs_ed("Pjkey")&"'>"
    response.write"<input name='nSerial' type='hidden' value ='"&rs_ed("Serial")&"'>"
    response.write"<input name='FMEANo' type='hidden' value ='"&rs_ed("FMEANo")&"'>"
    response.write"<input name='FMEARevision' type='hidden' value ='"&rs_ed("FMEARevision")&"'>"
    response.write"<td bgcolor='#F0F0F0'>"&trim(rs_ed("Serial"))&"-"&rs_ed("FMEANo")&"-"&rs_ed("FMEARevision")&"</td>"
	response.write"<td><input name='ProjectName' class='notline' size='25' value ='"&rs_ed("ProjectName")&"'></td>"
    response.write"<td><input name='ItemName' class='notline' size='25' value ='"&rs_ed("ItemName")&"'></td>"
    response.write"<td><input name='Model' class='notline' size='20' value ='"&rs_ed("Model")&"'></td>"
    response.write"<td><input name='Author' class='notline' size='8' value ='"&rs_ed("Author")&"'></td>"
    response.write"<td><input name='BurdenDept' class='notline' size='20' value ='"&rs_ed("BurdenDept")&"'></td>"
	response.write"<td><input name='Onduty' class='notline' size='20' value ='"&rs_ed("Onduty")&"'></td>"
    response.write"<td><input name='CreateDate' class='notline' size='15' value ='"&rs_ed("CreateDate")&"'></td>"
    response.write"<td><input name='OK' value='�޶�' type='submit' class='smallInput'>" 
    response.write"[<a href='javascript:history.back()'>����</a>]" 
    response.write"</td></tr>"   
    call closeDB()

elseif request("act") = "verup" then

    SqlStr_e = "SELECT * FROM t_Fmea_PJ where PjID ="&request("ID")
    Set rs_ed = Server.CreateObject("ADODB.Recordset")        
    rs_ed.open SqlStr_e,conn,1,1
    response.write"<tr bgcolor='#F0F0F0'><td></td><td></td>"
    response.write"<input name='PjID' type='hidden' value ='"&rs_ed("PjID")&"'>"
    response.write"<input name='nPjkey' type='hidden' value ='"&rs_ed("Pjkey")&"'>"
    response.write"<input name='nSerial' type='hidden' value ='"&rs_ed("Serial")&"'>"
    response.write"<td>"&rs_ed("Serial") &"-"
    response.write"<input name='FMEANo' type='hidden' value ='"&rs_ed("FMEANo")&"'>"
	response.write rs_ed("FMEANo")&"-"       
    response.write"<input name='FMEARevision' class='notline' size='5' value ='"&rs_ed("FMEARevision")&"'></td>"
	response.write"<td>"&rs_ed("ProjectName")&"</td>"
    response.write"<td>"&rs_ed("ItemName")&"</td>"
    response.write"<td>"&rs_ed("Model")&"</td>"
    response.write"<td>"&rs_ed("Author")&"</td>"
    response.write"<td>"&rs_ed("BurdenDept")&"</td>"
	response.write"<td>"&rs_ed("Onduty")&"</td>"
    response.write"<td>"&rs_ed("CreateDate")&"</td>"
    response.write"<input name='PublishPerson' type='hidden' value ='"&rs_ed("PublishPerson")&"'>"
    response.write"<td><input name='OK' value='����' type='submit' class='smallInput'>"
    response.write"[<a href='javascript:history.back()'>����</a>]" 
    response.write"</td></tr>"   
    call closeDB()

else
    call MainBody()

%>

<tr>
<td width="75%" colspan=13 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">�ص���ҳ</a>] [<a href="javascript:history.back()">����</a>]  
</td><td  width="50%" align="right">
<% call clspage() %> 
</table>
</td>
</tr>
</table>
<% 
call closeDB() 
end if 
%>
</form>

</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=6>"
	response.write "û�м�¼��"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
	PjKey = rs("PjKey")
    Serial = rs("Serial")

	response.write"<tr bgcolor='#FCFCFC'>"
    
    if isnull(rs("ApplyPerson")) then
        response.write"<td bgcolor ='FFB4CF'>&nbsp;<a href=Appr_Browse.asp?cat=2&ID="&rs("PjID")&">����δ��(1)</a></td>"
    elseif isnull(rs("AuditPerson")) then
        response.write"<td bgcolor ='FFB4CF'>&nbsp;<a href=Appr_Browse.asp?cat=2&ID="&rs("PjID")&">����δ��(2)</a></td>"
    elseif isnull(rs("ApprovePerson")) then
        response.write"<td bgcolor ='FFB4CF'>&nbsp;<a href=Appr_Browse.asp?cat=2&ID="&rs("PjID")&">����δ��(3)</a></td>"
    elseif isnull(rs("PublishPerson")) then
        response.write"<td bgcolor ='FFE7EF'>&nbsp;<a href=Appr_Browse.asp?cat=2&ID="&rs("PjID")&">����δ��(4)</a></td>"
    else
        response.write"<td>&nbsp;<a href=Appr_Browse.asp?cat=2&ID="&rs("PjID")&">��������</a></td>"
    end if

    response.write"<td bgcolor ='#AFF7FF'>&nbsp;<a href=../History/His_Pj_Browse.asp?PjID="&rs("PjID")&">Link</a></td>"   
	response.write"<td><a href=../ProcessMan/Process_Browse.asp?PjID="&rs("PjID")&"&PjKey="&rs("PjKey")&">" &Pjkey&"</a></td>"
	response.write"<td>"&rs("ProjectName")&"</td>"
    response.write"<td>"&rs("ItemName")&"</td>"
    response.write"<td>"&rs("Model")&"</td>"
    response.write"<td>"&rs("Author")&"</td>"
    response.write"<td>"&rs("BurdenDept")&"</td>"
	response.write"<td>"&rs("Onduty")&"</td>"
    response.write"<td>"&rs("CreateDate")&"</td>"

    if request.cookies("fmea.power") > "1" then         
         response.write"<td><a href=Pj_Browse.asp?act=edit&ID="&rs("PjID")&">�޶�</a>&nbsp;"
         response.write"&nbsp;<a href=Pj_Browse.asp?act=verup&ID="&rs("PjID")&">����</a>&nbsp;"
         response.write"&nbsp;<a href=Pj_Browse.asp?act=del&ID="&rs("PjID")&">ɾ��</a></td>"
    end if
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=Pj_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Pj_Browse.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=Pj_Browse.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
Response.Write "<a href=Pj_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=Pj_Browse.asp?page=1>��ҳ</a>&nbsp;"
Response.Write "<a href=Pj_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Pj_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Pj_Browse.asp?page="&pagecount&">βҳ</a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub

%>

