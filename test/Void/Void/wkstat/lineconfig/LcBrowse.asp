<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "��ѯ" 	                
            select case Request("cata")
              case "�߼���ѯ"
              case "ģ����ѯ"
                    Response.Cookies("ca_Search")= Request("Search")
	            Response.Redirect "Lcbrowse.asp?Cat=0&Sc="&Request("Search")       		   		
              case "��Line��ѯ"
                    Response.Cookies("ca_Search")= Request("Search")
                    Response.Redirect "Lcbrowse.asp?Cat=1&Sc="&Request("Search")
              case "�����ֲ�ѯ"
                    Response.Cookies("ca_Search")= Request("Search")
                    Response.Redirect "Lcbrowse.asp?Cat=2&Sc="&Request("Search")
              case "�����̲�ѯ"
                    Response.Cookies("ca_Search")= Request("Search")
                    Response.Redirect "Lcbrowse.asp?Cat=3&Sc="&Request("Search")
              case "����λ��ѯ"
                    Response.Cookies("ca_Search")= Request("Search")
                    Response.Redirect "Lcbrowse.asp?Cat=4&Sc="&Request("Search")
              case "�����¼1"
                    Response.Redirect "Lcbrowse.asp?Cat=5&Sc="&Request("Search")		
          end select            
End Select

pgs =500
select case request("cat")
       case "" 
       sql="SELECT * FROM  v_wmaster where ��Ч ='Y' order by line"
       pgs =25
       case 0
       sortl = request("sortl")
       select case sortl
       case ""
       sql="SELECT * FROM v_wmaster where  line like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' Or ��λ like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' order by ����,��λ"
        case "line"
           sql="SELECT * FROM v_wmaster where line like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ��λ like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' order by line"
        case "model"
           sql="SELECT * FROM v_wmaster where line like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ��λ like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' order by ����"
         case "wogroup"
           sql="SELECT * FROM v_wmaster where line like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ��λ like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' order by ����"
         case "workstation"
           sql="SELECT * FROM v_wmaster where line like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ��λ like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' Or ���� like '%"&request.Cookies("ca_Search")& _
           "%' order by ��λ"
        end select
        response.Cookies("c_Sc")= request("Sc")
       case 1
         sql="SELECT * FROM v_wmaster where line = '"&request("Sc")&"' order by ����"
       case 2
         sql="SELECT * FROM v_wmaster where ���� = '"&request("Sc")&"' order by line"
       case 3
         sql="SELECT * FROM v_wmaster where ���� = '"&request("Sc")&"' order by line"  
       case 4
         sql="SELECT * FROM v_wmaster where  ��λ = '"&request("Sc")&"' order by line" 
       case 5
         sql="SELECT * FROM v_wmaster WHERE (���� IS NULL)"
       case 6
         sql="SELECT * FROM v_wmaster where  ��� = '"&request.Cookies("leader")&"' order by ��λ" 
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��λ��Ա����</title>
</head>

<body>
<form action="Lcbrowse.asp">
<table  width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=11>
<b>���̰���һ��</b></td>
</tr> 
<tr><td bgcolor="#E6F7FF" colspan=11>
<table width="100%" border="0" cellpadding="0" ><tr>
<td height="20" width ="15%" align='center' bgcolor="#e6f7ff" >
<a href="Lc_count.asp">��LINE��λ��</a></td>
<td height="20" width ="15%"  align='center' bgcolor="#e6f7ff" >
[<%=request.Cookies("leader") %>]<a href="Lcbrowse_ld.asp">�ҵĹ���Χ</a></td>
<td align='right'  width ="70%"  bgcolor="#e6f7ff" >
<select name='cata' class="smallinput">
<option>ģ����ѯ</option>
<option>��Line��ѯ</option>
<option>�����ֲ�ѯ</option>
<option>�����̲�ѯ</option>
<option>����λ��ѯ</option>
<option>�����¼1</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="��ѯ" name="OK" class="smallInput">
</td></tr></table>
</td></tr>


<% if request.Cookies("ca_Search") ="" then %>
<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="25">Line</td>
<td width="20%">����</td>
<td width="5%">����</td>
<td width="8%">��λ</td>
<td width="5%">����</td>
<td width="5%">����</td>
<td width="8%">ְ��</td>
<td width="8%">������</td>
<td width="8%">��Ч</td>
<td width="20%" align="center">�������</td>
</tr>
<% else %> 
<tr bgcolor="#FCFCFC" align="center">
<td width="5%" height="25"><a href=lcbrowse.asp?cat=0&sortl=line>Line</a></td>
<td width="20%"><a href=lcbrowse.asp?cat=0&sortl=model>����</a></td>
<td width="5%"><a href=lcbrowse.asp?cat=0&sortl=wogroup>����</a></td>
<td width="8%"><a href=lcbrowse.asp?cat=0&sortl=workstation>��λ</a></td>
<td width="5%">����</td>
<td width="5%">����</td>
<td width="8%">ְ��</td>
<td width="8%">������</td>
<td width="8%">��Ч</td>
<td width="20%" align="center">�������</td>
</tr>
<% end if %>

<% call MainBody() %>

<tr>
<td width="100%" colspan=11 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
<% call clspage() %>
</td><td  width="50%" align="right">
[<a href="javascript:history.back()">����</a>]   
<% Call CloseDB() %>
</table>
</td>
</tr>
</table>
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
	response.write"<tr bgcolor='#FCFCFC'>"	
        response.write"<td>"
        response.write "<a href=LcRecord.asp?line="&rs("line")&"&model="&replace(rs("����"),"","%")&"&wogroup="&rs("����")&"&workstation="&rs("��λ")&">"&rs("Line")&"</a></td>"
    
        response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("��λ")&"</td>"
	response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("����")&"</td>"
        response.write"<td>"&rs("ְ��")&"</td>"
        response.write"<td>"&rs("������")&"</td>"
        response.write"<td>"&rs("��Ч")&"</td>"
	response.write"<td>&nbsp;&nbsp;"
       ' response.write"<a href=LcEdit.asp?lineID="&rs("lineid")&">"&"ԭ��ҵ�߱��</a>&nbsp;&nbsp;"
       ' if hour(now()) < 10 then   
       ' response.write"<a href=LcAdd.asp?lineID="&rs("lineid")&">"&"��λ����</a>&nbsp;&nbsp;"
       ' else
       ' response.write"�ϰ�,��10����"
       ' end if
	response.write"</td></tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=Lcbrowse.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse.asp?page="&pagecount&">βҳ<a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=Lcbrowse.asp?page=1>��ҳ<a>&nbsp;&nbsp;"
Response.Write "<a href=Lcbrowse.asp?page="&page-1&">��ҳ<a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=Lcbrowse.asp?page=1>��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse.asp?page="&page-1&">��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=Lcbrowse.asp?page="&pagecount&">βҳ<a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����Ч��λ"
end sub


%>