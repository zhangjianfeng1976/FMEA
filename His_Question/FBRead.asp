<!--#include file="../includes/db.asp"-->
<% 
pgx=500
select case request("cat")	
    case 1
        sql="SELECT * FROM t_FailBase where ID = "&request("ID")		
    case 3
        sql="SELECT * FROM v_Read_FailBase where ItmID_fb = "&request("ItmID")	
    case del
        sql = "Delete FROM t_Failbase_Connect where ConnID="&request("ConnID")
        Call openDB()
        conn.execute(sql)
        Call CloseDB()
        Response.Write "<script>window.opener=null;window.close();</script>"
        Response.end		
end select

Call openDB()
rs.open sql,conn,1,1
If not rs.eof then
    rs.pagesize = pgx
    pagecount=rs.PageCount 
    page=int(request.QueryString ("page"))
    if page<=0 then page=1
    if request.QueryString("page")="" then page=1
    rs.AbsolutePage=page 
End if
 %>  
<html>
<head>
<title>��ȥ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Includes/Style.css" rel="stylesheet" type="text/css">
<script>
function showCustomer(str)
{
var xmlhttp;
if (str=="")
? {
? document.getElementById("txtHint").innerHTML="";
? return;
? }
if (window.XMLHttpRequest)
? {// code for IE7+, Firefox, Chrome, Opera, Safari
? xmlhttp=new XMLHttpRequest();
? }
else
? {// code for IE6, IE5
? xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
? }
xmlhttp.onreadystatechange=function()
? {
? if (xmlhttp.readyState==4 && xmlhttp.status==200)
??? {
??? document.getElementById("txtHint").innerHTML=xmlhttp.responseText;
??? }
? }
xmlhttp.open("GET","getcustomer.html?q="+str,true);
xmlhttp.send();
}
</script>
</head>

<body>

<% 
  if request("cat") = 1 then
    sql_h ="SELECT * FROM v_Read_FailBase where ID = "&request("ID")&" order by OthCheckTime"
	set rs_h = server.createobject("ADODB.Recordset")
    rs_h.open sql_h,conn,1,1
    response.write"<table width='60%' border='0' cellpadding='0' cellspacing='1' bgcolor='#5A8BCE' align='center' >"
    response.write"<tr bgcolor='#AAAAAA'>"
    response.write"<td colspan ='2' heigh ='30' align='center'><font size = '4' color ='white'>��ȥ�������ϸ����</font></td>"
    response.write"</tr>"
  elseif request("cat") = 3 then 
    response.write"<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#5A8BCE' align='center' >"
  end if
  call mainbody
 %>
</table>
<% call closeDB() %>
</body>
</html>

<%
Sub Mainbody()

if rs.eof then
  response.write "<tr><td width=""75%"" colspan=12>"
  response.write "û�����ݡ�"
  response.write "</td></tr>"
end if

n=1 
While Not rs.eof and n<=rs.pagesize 
  response.write"<tr bgcolor='#FCFCFC'>"
  response.write"<td width='10%'>����</td><td>"&rs("Model")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>����</td><td>"&rs("Dtae")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>��������</td><td>"&rs("LocationCata")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>����</td><td>"&rs("MDepartment")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>��������</td><td>"&rs("FailContent")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>��ҵ����</td><td>"&rs("WorkTheme")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>��Ʒ����</td><td>"&rs("PartsType")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>��ҵ����</td><td>"&rs("WorkCata")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>�������</td><td>"&rs("FailCatalog")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>ԭ��</td><td>"&rs("FailReason")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>�Բ�</td><td>"&rs("FailAction")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>��Ӧ����</td><td>"&rs("ActionMode")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>FMEA���ַ���</td><td>"&rs("FMEARsCata")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  response.write"<td>FMEA��������</td><td>"&rs("FMEAResearch")&"</td>"
  response.write"</tr><tr bgcolor='#FCFCFC'>"
  if request("cat") = 3 then 
    response.write"<td colspan ='2' align ='center'>��<a href = 'FBRead.asp?act=del&ConnID="&rs("ConnID")&"'>ɾ��</a>��</td>"
  end if
  response.write"</tr>"
  n=n+1 
  rs.movenext '��ʾҳ�������' 
Wend 

if request("cat") = 1 then 
   response.write"<tr bgcolor='#FCFCFC'>"
   response.write"<td colspan ='2' align ='center' height ='20'>��<a href = 'javascript:history.back()'>����</a>��</td>"
   response.write"</tr>"
   response.write"<tr bgcolor='#FCFCFC'><td colspan ='2'>"
   While Not rs_h.eof
	  response.write"<div><a href=../ProcessMan/Process_Browse_total.asp?cat=5&ItmID="&rs_h("ItmID_fb")&">"
      response.write rs_h("OthCheckTime")&"&nbsp&nbsp"&rs_h("OthChecker")
	  response.write"&nbsp&nbspˮƽչ����&nbsp&nbsp"&rs_h("Pjkey")&"</a></div>"	  
	  rs_h.movenext
   Wend
   response.write"</td></tr>"
End If   

End Sub
%>

 
