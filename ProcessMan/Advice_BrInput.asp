<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "����" 	                
	     Response.Redirect "Advice_BrInput.asp?Cat=0&Sc="&Request("Search")
       Case "ȫ��������" 	                
	     Response.Redirect "Advice_BrInput.asp?Cat=1&Sc="&Request("Search")	             
End Select

pgs =2500
select case request("cat")
       case "" 
         ttl = request.cookies("fmea.Pjkey")
         sql="SELECT * FROM v_Fmea_total where Pjkey = '"&request.cookies("fmea.Pjkey")&"' and RPN > 2 and  AdID is null order by PageNo,WorkQueue"
       case 0
          ttl = request.cookies("fmea.Pjkey")
          sql = "SELECT * FROM v_Fmea_total where "
          sql = sql & "(AdvContent like '%" & request("sc") & "%' or "
          sql = sql & "Rspnser like '%" & request("sc") & "%' or "
          sql = sql & "ActionContent like '%" & request("sc") & "%'"
          sql = sql & ") and Pjkey = '"&request.cookies("fmea.Pjkey")&"' and  AdID is null"
       case 1
          ttl = "ȫ��"
          sql = "SELECT * FROM v_Fmea_total where "
          sql = sql & "(AdvContent like '%" & request("sc") & "%' or "
          sql = sql & "Rspnser like '%" & request("sc") & "%' or "
          sql = sql & "ActionContent like '%" & request("sc") & "%'"
          sql = sql & ") order by ItmID"
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

<!--Java Script Start
var winconf = "menubar=no,toolbar=yes,location=no,directories=no,status=no,scrollbar=no,resizeable=no,";
function about(filename,w,h) 
{
about = window.open(filename,"Copyright",winconf+"width="+w+",height="+h);
about.focus();
}
//Java Script End-->
</script>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ָ��ƴ�ʩ��Ŀһ��</title>
</head>

<body bgcolor ="#FCFCFC">
<form action="Advice_BrInput.asp">
<!--<div style ="float:right ;width��200px;hight:200px;">
<input name="Search" type="text" class="smallInput" size="40">
<input type="submit" value="����" name="OK" class="smallInput">
<input type="submit" value="ȫ��������" name="OK" class="smallInput">
&nbsp;&nbsp;
</div>-->

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=20>
<b><% = ttl %>δ������Ŀһ��</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="2%" height="20" rowspan ="2">˳��</td>
<td width="8%" rowspan ="2">FMEA��ż��汾</td>
<td width="8%" rowspan ="2">��������</td>
<td width="5%" rowspan ="2">��ҵ��׼��ҳ��</td>
<td width="8%" rowspan ="2">��Ʒ���</td>
<td width="2%" rowspan ="2">Sev</td>
<td width="2%" rowspan ="2">Occ</td>
<td width="2%" rowspan ="2">Det</td>
<td width="2%" rowspan ="2">RI</td>
<td width="12%" rowspan ="2">�����ʩ</td>
<td width="5%" rowspan ="2">������</td>
<td width="8%" rowspan ="2">Ŀ�������</td>
<td width="15%" colspan ="5">��ʩ���</td>
<%
if request.cookies("fmea.power") > "1" then
   response.write "<td width='5%' rowspan ='2'>����</td>"
end if
%>
</tr><tr bgcolor="#FCFCFC" align="center">
<td width="10%">���ô�ʩ</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RI</td>
</tr>

<% 
call MainBody()
 %>

<tr>
<td width="75%" colspan=20 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">�ص���ҳ</a>][<a href="Process_Browse.asp">����һ��</a>][<a href="javascript:history.back()">����</a>]
</td><td  width="50%" align="right">
<% call clspage() %>   
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
	ItmID = rs("ItmID")
	response.write"<tr bgcolor='#FCFCFC'>"
    response.write"<td>"&rs("WorkQueue")&"</td>"
    response.write"<td>"&rs("PjKey")&"</td>"
    response.write"<td>"&rs("SubPjName")&"</td>"	
    response.write"<td>"&rs("PageNo")&"</td>"
    response.write"<td>"&rs("PartsNo")&"</td>"
        
    response.write"<td>"&rs("Sev")&"</td>"
    response.write"<td>"&rs("Occ")&"</td>"
    response.write"<td>"&rs("Det")&"</td>"
    response.write"<td>"&rs("RPN")&"</td>" 
	
    response.write"<td>"&rs("AdvContent")&"</td>"	
	response.write"<td>"&rs("Rspnser")&"</td>"
    response.write"<td>"&rs("FshDate")&"</td>"
    response.write"<td>"&rs("ActionContent")&"</td>"
    response.write"<td>"&rs("ResultS")&"</td>"
    response.write"<td>"&rs("ResultO")&"</td>"
    response.write"<td>"&rs("ResultD")&"</td>"
    response.write"<td>"&rs("ResultRPN")&"</td>" 
    if request.cookies("fmea.power") > "1" then
        response.write"<td>&nbsp;<a href=Advice_Add.asp?ItmID="&ItmID&">��ʩ����</a></td>"
    end if
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=Advice_BrInput.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Advice_BrInput.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=Advice_BrInput.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
Response.Write "<a href=Advice_BrInput.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=Advice_BrInput.asp?page=1>��ҳ</a>&nbsp;"
Response.Write "<a href=Advice_BrInput.asp?page="&page-1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Advice_BrInput.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Advice_BrInput.asp?page="&pagecount&">βҳ</a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub


%>

