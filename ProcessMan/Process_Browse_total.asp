<!--#include file="../includes/db.asp"-->
<%
select case request("cat")
    case "" 
    case 0
       sql="SELECT * FROM v_Fmea_total where (SubPjName like '%"&request("Sc")& _
	       "%' Or WorkTheme like '%"&request("Sc")& _
           "%' Or PageNo like '%"&request("Sc")& _
           "%' Or PartsType like '%"&request("Sc")& _
           "%' Or PartsNo like '%"&request("Sc")& _
           "%' Or PartsName like '%"&request("Sc")& _
           "%' Or WorkCata like '%"&request("Sc")& _
           "%' Or WorkContent like '%"&request("Sc")& _
           "%' Or NGMode like '%"&request("Sc")& _
           "%' Or NGEffect like '%"&request("Sc")& _
           "%' Or ProDesign like '%"&request("Sc")& _
           "%') ORDER BY SubPjName,PageNo" 
     case 5
         sql="SELECT * FROM v_Fmea_total where ItmID ="&request("ItmID")&" ORDER BY SubPjName,PageNo,WorkQueue"   
     case 6
         sql="SELECT * FROM v_Fmea_total where ItmID IN (Select ItmID_Fb from t_failbase_connect Where Fb_His_ID ="&request("ID")&") ORDER BY SubPjName,PageNo,WorkQueue"
     case 7
         sql="SELECT * FROM v_Fmea_total where safeparts is not null and Sev <> 4 ORDER By SubPjName,PageNo,WorkQueue"	 
end select

call openDB()
pgs =2500
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

<Script LANGUAGE = "JAVASCRIPT" type="text/javascript">
<!--Java Script Start
var winconf = "menubar=no,toolbar=no,location=no,directories=no,status=no,scrollbar=no,resizeable=no,";
function about(filename,w,h) 
{
about = window.open(filename,"Copyright",winconf+"left=0,width="+w+",height="+h);
about.focus();
}
//Java Script End-->
</SCRIPT> 


<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE2 {color: #0000FF}
-->
</style>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ֹ�����Ŀһ��</title>
</head>

<body bgcolor ="#FCFCFC">
<form action="Process_Browse_Total.asp">
<input type='hidden' value="<%=request("Sc") %>" name='sc'>
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">

<tr><td width="75%" colspan=18 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">�ص���ҳ</a>] [<a href="javascript:history.back()">����</a>] 
</td><td  width="50%" align="right">
 <% call clspage() %> 
</td></tr>
</table>
</td></tr>

<tr bgcolor="#E6F7FF"><td height="30" align="center" colspan=18><b>ȫ������ҵ���������б�</b></td></tr> 

<tr bgcolor="#FCFCFC" align="center">
<td height="20" colspan = "4">FMEA��ż��汾</td>
<td height="20" colspan = "2">��Ŀ</td> 
<td height="20" colspan = "2">����</td> 
<td height="20" colspan = "2">����</td>
<td height="20" colspan = "5"></td>
</tr>
<tr bgcolor="#FCFCFC" align="center">
<td width="2%" height="20">˳��</td>
<td width="5%">��������</td>
<td width="3%">��׼��ҳ��</td>
<td width="2%">��ҵ����</td>
<td width="7%">��Ʒ���</td>
<td width="8%">��Ʒ���Ƽ�����</td>
<td width="8%">��ҵ����������</td>
<td width="10%">Ǳ��ȱ��ģʽ</td>
<td width="10%">Ǳ��ȱ�ݺ��</td>
<td width="7%">���й�����Ƽ����̿���</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RI</td>
<td width="4%">��ʩ</td>
</tr>

<% 
call MainBody() 
%>

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
  PjKey = rs("PjKey")
  PjID  = rs("PjID")
  response.write"<tr bgcolor='#F0F0F0'>"
  response.write"<td  colspan = '4'>"&PjKey&"</td>"
  response.write"<td  colspan = '2'>"&rs("ProjectName")&"</td>"
  response.write"<td  colspan = '2'>"&rs("Model")&"</td>"
  response.write"<td  colspan = '2'>"&rs("ItemName")&"</td>"
  response.write"<td  colspan = '5'></td>"
  response.write"</tr>"  
  response.write"<tr bgcolor='#FCFCFC'>"
        
  if rs("ItmStatus") = "N" then
    response.write"<td bgcolor ='FFFF00'>"&rs("WorkQueue")&"</td>"
    response.write"<td bgcolor ='FFFF00'>"&rs("SubPjName")&"</td>"
    response.write"<td bgcolor ='FFFF00'><div>(��ҵ��ֹ)</div>"&rs("PageNo")&"</td>"
  else
    response.write"<td>"&rs("WorkQueue")&"</td>"
    response.write"<td>"&rs("SubPjName")&"</td>"
    response.write"<td>"&rs("PageNo")&"</td>"
  end if
   
  response.write"<td>"&rs("WorkTheme")&"</td>"
	
  if isnull(rs("PartsOK")) and isnull(rs("SafeParts"))  then
    response.write"<td>"
    response.write rs("PartsNo")
    response.write"<a href = http://km-square.km.local/kmc-portal/mp/pgm/ItemSearch_c/Pages/ITEM_Search_P_c.aspx?ITEMNO="
    response.write rs("PartsNo")&"&p_QASTATUS=%20 target ='_blank'>"
    response.write"<img border='0' src='../images/drawing.png' width='20' height='20' align='left'>" 
    response.write"</a></td>"
  elseif not isnull(rs("PartsOK")) then
    response.write"<td bgcolor ='FF9955'>"
    response.write rs("PartsNo")
    response.write"<a href = http://km-square.km.local/kmc-portal/mp/pgm/ItemSearch_c/Pages/ITEM_Search_P_c.aspx?ITEMNO="
    response.write rs("PartsNo")&"&p_QASTATUS=%20 target ='_blank'>"
    response.write"<img border='0' src='../images/drawing1.png' width='20' height='20' align='left'></a>"
    response.write"<a href = '../His_Question/PartsFamily.asp?cat=0&sc="&trim(rs("PartsNo"))&"'>"
    response.write"<img border='0' src='../images/family.png' width='20' height='20' align='left'></a>"
    response.write"</td>"
  elseif not isnull(rs("SafeParts")) then
    response.write"<td>"
    response.write rs("PartsNo")
    response.write"<a href = http://km-square.km.local/kmc-portal/mp/pgm/ItemSearch_c/Pages/ITEM_Search_P_c.aspx?ITEMNO="
    response.write rs("PartsNo")&"&p_QASTATUS=%20 target ='_blank'>"
    response.write"<img border='0' src='../images/drawing.png' width='20' height='20' align='left'></a>"
    response.write"<img border='0' src='../images/safep.png' width='20' height='20' align='left'></a>"
    response.write"<img border='0' src='../images/sign.gif'  align='left'></a>"
    response.write"</td>"
  end if

  response.write"<td><div>("&rs("PartsType")&")</div><div>"&rs("PartsName")&"</div></td>"
  response.write"<td><div>("&rs("WorkCata")&")</div><div>"&rs("WorkContent")&"</div></td>" 
  response.write"<td>"&rs("NGMode")&"</td>"
  response.write"<td>"&rs("NGEffect")&"</td>"
  response.write"<td>"&rs("ProDesign")&"</td>"
  response.write"<td align ='center'>"&rs("Sev")&"</td>"
  response.write"<td align ='center'>"&rs("Occ")&"</td>"
  response.write"<td align ='center'>"&rs("Det")&"</td>"
        
  vRPN = trim(rs("RPN"))
  if vRPN > 2 Then
    response.write"<td bgcolor = '#DDFFAA'><span class='STYLE1'>&nbsp;"& vRPN &"</span></td>"    
    response.write"<td bgcolor = '#DDFFAA'>"
    'response.write"<div><a href=Advice_Browse.asp?cat=2&ItmID="&ItmID&">��Ӧ</a>&nbsp;"
    'response.write"<a href=Process_Browse_Total.asp?cat=0&sc="&request("sc")&" onClick=about('Advice_Read.asp?cat=2&ItmID="&ItmID&"',850,250)>�鿴</a></div>"
    response.write"<a href=# onClick=about('Advice_Read.asp?cat=2&ItmID="&ItmID&"',850,250)>�鿴</a>&nbsp;&nbsp;"
	if request.cookies("fmea.power") > "1" then  response.write"<a href=Process_Browse.asp?Act=Edit&PjID="&PjID&"&PjKey="&PjKey&"&ItmID="&ItmID&">�޶�</a>"
    if not isnull(rs("ItmID_his")) then response.write"<div><a href=../history/His_Process_Browse.asp?ItmID="&ItmID&">�޶�����</a></div>"
    if not isnull(rs("ItmID_fb")) then response.write"<div><a href= # onClick=about('../His_Question/FBRead.asp?cat=3&ItmID="&ItmID&"',600,600)>��ȥ����</a></div>"
    response.write"</td>"        
  else
    response.write"<td>"& vRPN &"</td>" 
    response.write"<td>"
    if request.cookies("fmea.power") > "1" then  response.write"<a href=Process_Browse.asp?Act=Edit&PjID="&PjID&"&PjKey="&PjKey&"&ItmID="&ItmID&">�޶�</a>"
    if not isnull(rs("ItmID_his")) then response.write"<div><a href=../history/His_Process_Browse.asp?ItmID="&ItmID&">�޶�����</a></div>"
    if not isnull(rs("ItmID_fb")) then response.write"<div><a href= # onClick=about('../His_Question/FBRead.asp?cat=3&ItmID="&ItmID&"',600,600)>��ȥ����</a></div>"
    response.write"</td>"
  end if 

  response.write"</tr>"     
  rs.MoveNext
n=n+1
wend
End Sub


Sub clspage()

If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
  Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
  Response.Write "<a href=Process_Browse_total.asp?page="&page+1&">��ҳ</a>&nbsp;"
  Response.Write "<a href=Process_Browse_total.asp?page="&pagecount&">βҳ</a>"
  
elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
  Response.Write "<a href=Process_Browse_total.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
  Response.Write "<a href=Process_Browse_total.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
  Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
  Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
  Response.Write "<a href=Process_Browse_total.asp?page=1>��ҳ</a>&nbsp;"
  Response.Write "<a href=Process_Browse_total.asp?page="&page-1&">��ҳ</a>&nbsp;"
  Response.Write "<a href=Process_Browse_total.asp?page="&page+1&">��ҳ</a>&nbsp;"
  Response.Write "<a href=Process_Browse_total.asp?page="&pagecount&">βҳ</a>"
end if 

Response.Write "&nbsp;��<input type='Text' name='pagename' size='2' class='smallInput'  value ='"&page& "'>ҳ"
Response.Write "<input type='submit' value='GO' name='ok' class='smallInput'>"
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"

end sub
%>

