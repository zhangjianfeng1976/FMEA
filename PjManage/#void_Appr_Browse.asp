<!--#include file="../includes/db.asp"-->
<%
if  Request.cookies("fmea.truename") = ""  then
     Response.Write "����[<a href ='../login.asp' target ='_parent'>��¼</a>]����в�����ѡ��[<a href='javascript:history.back()'>����</a>] "
     Response.End
end if

if (request.cookies("fmea.Pjkey") ="" or  request.cookies("fmea.Pjkey") <> request("Pjkey")) and  request("PjKey") <>"" then  
  response.cookies("fmea.Pjkey") = request("Pjkey")
end if 

if (request.cookies("fmea.PjID") ="" or  request.cookies("fmea.PjID") <> request("PjID")) and  request("PjID") <>"" then  
  response.cookies("fmea.PjID") = request("PjID")
end if 

Select Case Request("OK") 
       Case "����" 	                
	     Response.Redirect "Appr_Browse.asp?Cat=0&Sc="&Request("Search")
       Case "ֻ��ʾδ���" 	                
	     Response.Redirect "Appr_Browse.asp?Cat=1"			   	
End Select

Select Case Request("stmp") 
       Case 1 
          if  Request.cookies("fmea.power") < "3"  then
              Response.Write "��δ���㹻Ȩ������![<a href='javascript:history.back()'>����</a>] "
              Response.End
          end if
          SqlStr1 = "Update t_Fmea_PJ set "
          SqlStr1 = SqlStr1 & "ApplyPerson = '"& Request.Cookies("fmea.truename") & "',"
          SqlStr1 = SqlStr1 & "ApplyDate = '"& now() & "'"
          SqlStr1 = SqlStr1 & " Where PjID = '"& request("PjID") & "'"
          call openDB()
          conn.execute(SqlStr1)
          call closeDB()	                
	  Response.Redirect "Appr_Browse.asp?cat=2&ID="& request("PjID")
       Case 2
          if  Request.cookies("fmea.power") < "4"  then
              Response.Write "��δ���㹻Ȩ�����![<a href='javascript:history.back()'>����</a>] "
              Response.End
          end if
          SqlStr2 = "Update t_Fmea_PJ set "
          SqlStr2 = SqlStr2 & "AuditPerson = '"& Request.Cookies("fmea.truename") & "',"
          SqlStr2 = SqlStr2 & "AuditDate = '"& now() & "'"
          SqlStr2 = SqlStr2 & " Where PjID = '"& request("PjID") & "'"
          call openDB()
          conn.execute(SqlStr2)
          call closeDB()
          Response.Redirect "Appr_Browse.asp?cat=2&ID="& request("PjID")
       Case 3
          if  Request.cookies("fmea.power") < "5"  then
              Response.Write "��δ���㹻Ȩ�޳���![<a href='javascript:history.back()'>����</a>] "
              Response.End
          end if
          SqlStr3 = "Update t_Fmea_PJ set "
          SqlStr3 = SqlStr3 & "ApprovePerson = '"& Request.Cookies("fmea.truename") & "',"
          SqlStr3 = SqlStr3 & "ApproveDate = '"& now() & "'"
          SqlStr3 = SqlStr3 & " Where PjID = '"& request("PjID") & "'"
          call openDB()
          conn.execute(SqlStr3)
          call closeDB()
          Response.Redirect "Appr_Browse.asp?cat=2&ID="& request("PjID")
       Case 4
          if  Request.cookies("fmea.power") < "3"  then
              Response.Write "��δ���㹻Ȩ���䲼![<a href='javascript:history.back()'>����</a>] "
              Response.End
          end if
          SqlStr4 = "Update t_Fmea_PJ set "
          SqlStr4 = SqlStr4 & "PublishPerson = '"& Request.Cookies("fmea.truename") & "',"
          SqlStr4 = SqlStr4 & "PublishDate = '"& now() & "'"
          SqlStr4 = SqlStr4 & " Where PjID = '"& request("PjID") & "'"
          call openDB()
          conn.execute(SqlStr4)
          call closeDB()
          Response.Redirect "Appr_Browse.asp?cat=2&ID="& request("PjID")		   	
End Select

pgs =2500
select case request("cat")
       case "" 
         sql="SELECT * FROM t_Fmea_PJ WHERE serial ='"&request.cookies("fmea.serial")&"'"
         pgs =25
       case 0
         sql="SELECT * FROM t_Fmea_PJ where (PjKey like '%"&request("Sc")& _
           "%' Or ProjectName like '%"&request("Sc")& _
           "%' Or ItemName like '%"&request("Sc")& _
           "%' Or Model like '%"&request("Sc")& _
           "%' Or Author like '%"&request("Sc")& _
           "%' Or ApplyPerson like '%"&request("Sc")& _
           "%' Or AuditPerson like '%"&request("Sc")& _
           "%' Or ApprovePerson like '%"&request("Sc")& _
           "%' Or PublishPerson like '%"&request("Sc")& _
           "%') and serial ='"&request.cookies("fmea.serial")&"' ORDER BY PjID" 
        case 1
          sql="SELECT * FROM t_Fmea_PJ where " _
            & "ApplyPerson is null " _
            & "Or AuditPerson is null " _
            & "Or ApprovePerson is null " _
            & "Or PublishPerson is null " _
            & " ORDER BY PjID" 
        case 2
          sql="SELECT * FROM t_Fmea_PJ where PjID ="&request("ID")
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
<title>���Ϲ���</title>
</head>

<body bgcolor ="#FCFCFC"  marginheight="0" topmargin="0" bottommargin="0" leftmargin="0">
<form action="Appr_Browse.asp">
<!-- <table width="100%" bgcolor="#9AAADA" border="0" cellpadding="0" cellspacing="0">
<tr><td height="30" align="left">
<input name="Search" type="text" class="smallInput" size="40">
<input type="submit" value="����" name="OK" class="smallInput">
<input type="submit" value="ֻ��ʾδ���" name="OK" class="smallInput">
</td></tr></table> -->
&nbsp;
<table width="80%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align ="center">
<tr bgcolor="#E6F7FF"><td height="30" align="center" colspan="8"><b>���ֳ��������б�</b></td></tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="10%" height="20">FMEA��ż��汾</td>
<td width="10%">�޶�����</td>
<td width="15%">��Ŀ</td>
<td width="20%">����</td>
<td width="10%">����</td>
<td width="10%">���</td>
<td width="10%">����</td>
<td width="10%">�䲼����</td>
</tr>

<% call MainBody() %>

<tr>
<td width="80%" colspan=10 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">�ص���ҳ</a>][<a href="Pj_Browse.asp">����һ��</a>][<a href="javascript:history.back()">����</a>]  

</td><td  width="50%" align="right">
<!-- call clspage() --> 
</table>
</td>
</tr>
</table>
</form>
<% call closeDB() %>
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
	PjID = rs("PjID")
	response.write"<tr bgcolor='#FCFCFC' align ='center'>"
	response.write"<td height =20><a href=..\ProcessMan\Process_Browse.asp?PjID="&rs("PjID")&"&PjKey="&rs("PjKey")&">" &rs("Pjkey")&"</a></td>"
        response.write"<td height =20><a href=..\History\His_Pj_Browse.asp?PjID="&rs("PjID")&">Link</a></td>"
        response.write"<td>"&rs("ProjectName")&"</td>"
        response.write"<td>"&rs("ItemName")&"</td>"
        if isnull(rs("ApplyPerson")) then
        response.write"<td><a href=Appr_Browse.asp?stmp=1&PjID="&PjID&">ǩ��</a></td><td></td><td></td><td></td>"
        elseif isnull(rs("AuditPerson")) then
        response.write"<td>"&rs("ApplyPerson")&"</div><div>"&rs("ApplyDate")&"</div></td>"
        response.write"<td><a href=Appr_Browse.asp?stmp=2&PjID="&PjID&">ǩ��</a></td><td></td><td></td>"
        elseif isnull(rs("ApprovePerson")) then
        response.write"<td><div>"&rs("ApplyPerson")&"</div><div>"&rs("ApplyDate")&"</div></td>"
        response.write"<td><div>"&rs("AuditPerson")&"</div><div>"&rs("AuditDate")&"</div></td>"
        response.write"<td><a href=Appr_Browse.asp?stmp=3&PjID="&PjID&">ǩ��</a></td><td></td>"
        elseif isnull(rs("PublishPerson")) then
        response.write"<td><div>"&rs("ApplyPerson")&"</div><div>"&rs("ApplyDate")&"</div></td>"
        response.write"<td><div>"&rs("AuditPerson")&"</div><div>"&rs("AuditDate")&"</div></td>"
        response.write"<td><div>"&rs("ApprovePerson")&"</div><div>"&rs("ApproveDate")&"</div></td>"
        response.write"<td><a href=Appr_Browse.asp?stmp=4&PjID="&PjID&">ǩ��</a></td>"
        else
        response.write"<td><div>"&rs("ApplyPerson")&"</div><div>"&rs("ApplyDate")&"</div></td>"	
	response.write"<td><div>"&rs("AuditPerson")&"</div><div>"&rs("AuditDate")&"</div></td>"
        response.write"<td><div>"&rs("ApprovePerson")&"</div><div>"&rs("ApproveDate")&"</div></td>"
        response.write"<td><div>"&rs("PublishPerson")&"</div><div>"&rs("PublishDate")&"</div></td>" 
        end if   
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=Appr_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Appr_Browse.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=Appr_Browse.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
Response.Write "<a href=Appr_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=Appr_Browse.asp?page=1>��ҳ</a>&nbsp;"
Response.Write "<a href=Appr_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Appr_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Appr_Browse.asp?page="&pagecount&">βҳ</a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub


%>

