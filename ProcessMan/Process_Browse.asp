<!--#include file="../includes/db.asp"-->
<%
if (request.cookies("fmea.Pjkey") ="" or  request.cookies("fmea.Pjkey") <> request("Pjkey")) and  request("PjKey") <>"" then  
response.cookies("fmea.Pjkey") = request("Pjkey")
end if 

if (request.cookies("fmea.PjID") ="" or  request.cookies("fmea.PjID") <> request("PjID")) and  request("PjID") <>"" then  
response.cookies("fmea.PjID") = request("PjID")
end if 

Select Case Request("OK") 
       Case "ȫ��������" 	                
	     Response.Redirect "Process_Browse_total.asp?Cat=0&Sc="&Request("Search")	

       Case "����"
            If Request("Vle") ="Check" Then
                If Request("Search") ="" then
                   Response.Write "<script language='JavaScript' type='text/javascript'>alert('��������Ҫ��ѯ��RIֵ');</script>"  
                Else
                   Response.Redirect "Process_Browse.asp?Cat=6&Sc="&Request("Search")
                End If
            Else 	                
	           Response.Redirect "Process_Browse.asp?Cat=1&Sc="&Request("Search")
            End if

       Case "ȫϵ������" 	                
	     Response.Redirect "Process_Browse.asp?Cat=3&Sc="&Request("Search")
       
       Case "���Ӧ��Ŀ" 	                
	     Response.Redirect "Process_Browse.asp?Cat=2"

       Case "���Ʋ�Ʒ" 	                
	     Response.Redirect "Process_Browse.asp?Cat=4"

       Case "��������"
         Response.Redirect "../PjManage/Appr_Browse.asp?cat=2&ID="&request.cookies("fmea.PjID")
     
       Case "����¼��"
         Response.Redirect "Process_Add.asp"
 
       Case "��ʩ¼��"
             Response.Redirect "Advice_BrInput.asp"  

       Case "������"
             if request("page") ="" then
               Response.Redirect "Advice_Browse.asp?page=1"
             else
               Response.Redirect "Advice_Browse.asp?page="&request("page")      
             end if

       Case "����"
             Sev = int(request("Sev"))
             Occ = int(request("Occ"))
             Det = int(request("Det"))
             RPN = round((Sev * Occ * Det)^(1/3),1)
             
             If request("oRPN")<2.1 and RPN > 2 then
                Response.Write "���޶���������Ҫ���棬��ѡ�㡾�޶�(����)��"
                Response.Write "&nbsp;&nbsp;&nbsp;��<a href='javascript:history.go(-1)'>����</a>��<Br><Br>" 
                Response.End  
             End If
                 
            SqlStr_c ="Update t_Fmea_Item set "
            SqlStr_c = SqlStr_c & "WorkQueue = '"& request("WorkQueue") & "',"
            SqlStr_c = SqlStr_c & "SubPjName = '"& request("SubPjName") & "',"
            SqlStr_c = SqlStr_c & "PageNo = '"& request("PageNo") & "',"  
            SqlStr_c = SqlStr_c & "WorkTheme = '"& request("WorkTheme") & "',"			
            SqlStr_c = SqlStr_c & "PartsNo = '"& request("PartsNo") & "',"
            SqlStr_c = SqlStr_c & "PartsName = '"& request("PartsName") & "',"
            SqlStr_c = SqlStr_c & "PartsType = '"& request("PartsType") & "',"
            SqlStr_c = SqlStr_c & "WorkCata = '"& request("WorkCata") & "',"
            SqlStr_c = SqlStr_c & "WorkContent = '"& request("WorkContent") & "',"
            SqlStr_c = SqlStr_c & "NGMode = '"& request("NGMode") & "',"
            SqlStr_c = SqlStr_c & "NGEffect = '"& request("NGEffect") & "',"
            SqlStr_c = SqlStr_c & "ProDesign = '"& request("ProDesign") & "',"
            SqlStr_c = SqlStr_c & "Sev = '"& request("Sev") & "',"
            SqlStr_c = SqlStr_c & "Occ = '"& request("Occ") & "',"
            SqlStr_c = SqlStr_c & "Det = '"& request("Det") & "',"
            SqlStr_c = SqlStr_c & "RPN = '"& RPN & "'"
            SqlStr_c = SqlStr_c & " Where ItmID = '"& request("ItmID") & "'"
            call openDB()
            conn.execute(SqlStr_c)
            call closeDB()

            Response.Redirect "Process_Browse.asp?page="&request("page")

         Case "�޶�(����)" 
            
            sql = "Insert into t_Fmea_Item_history "
            sql = sql & "(ItmID,PjID,Pjkey,SubPjName,PageNo,WorkQueue,WorkTheme,PartsNo,PartsName,PartsType,WorkCata,"
            sql = sql & "WorkContent,NGMode,NGEffect,ProDesign,Sev,Occ,Det,RPN,ItmStatus,InUser,InTime)"
            sql = sql & " Select ItmID, PjID ,Pjkey,SubPjName,PageNo,WorkQueue,WorkTheme,PartsNo,"
            sql = sql & "PartsName,PartsType,WorkCata,WorkContent,NGMode,NGEffect,"
            sql = sql & "ProDesign,Sev,Occ,Det,RPN,'N',InUser,InTime"
            sql = sql & " from t_Fmea_Item Where ItmID = '"& request("ItmID") & "'"
            
            Sev = int(request("Sev"))
            Occ = int(request("Occ"))
            Det = int(request("Det"))
            RPN = round((Sev * Occ * Det)^(1/3),1)
            SqlStr_c ="Update t_Fmea_Item set "
            SqlStr_c = SqlStr_c & "WorkQueue = '"& request("WorkQueue") & "',"
            SqlStr_c = SqlStr_c & "SubPjName = '"& request("SubPjName") & "',"
            SqlStr_c = SqlStr_c & "PageNo = '"& request("PageNo") & "'," 
			SqlStr_c = SqlStr_c & "WorkTheme = '"& request("WorkTheme") & "',"
            SqlStr_c = SqlStr_c & "PartsNo = '"& request("PartsNo") & "',"
            SqlStr_c = SqlStr_c & "PartsName = '"& request("PartsName") & "',"
            SqlStr_c = SqlStr_c & "PartsType = '"& request("PartsType") & "',"
            SqlStr_c = SqlStr_c & "WorkCata = '"& request("WorkCata") & "',"
            SqlStr_c = SqlStr_c & "WorkContent = '"& request("WorkContent") & "',"
            SqlStr_c = SqlStr_c & "NGMode = '"& request("NGMode") & "',"
            SqlStr_c = SqlStr_c & "NGEffect = '"& request("NGEffect") & "',"
            SqlStr_c = SqlStr_c & "ProDesign = '"& request("ProDesign") & "',"
            SqlStr_c = SqlStr_c & "Sev = '"& request("Sev") & "',"
            SqlStr_c = SqlStr_c & "Occ = '"& request("Occ") & "',"
            SqlStr_c = SqlStr_c & "Det = '"& request("Det") & "',"
            SqlStr_c = SqlStr_c & "RPN = '"& RPN & "',"
            SqlStr_c = SqlStr_c & "InUser = '"& request.cookies("fmea.truename") & "',"
            SqlStr_c = SqlStr_c & "InTime = '"& now() & "'"
            SqlStr_c = SqlStr_c & " Where ItmID = '"& request("ItmID") & "'"
            
            call openDB()
            conn.BeginTrans
            conn.execute(sql)
            conn.execute(SqlStr_c)           
            conn.CommitTrans
            call closeDB()

            Response.Redirect "Process_Browse.asp?page="&request("page")
          
           Case "��ֹ" 
            Response.Write "�㽫Ҫ��ֹ���ݣ���ؼ�¼����ʾ��ҵ��ֹ��Ҫ�����𣿣������������ع������أ�<Br><Br>"
            Response.Write "��<a href='Process_del.asp?cat=1&ItmID="&request("ItmID")&"'>û���⣬���ְɣ�</a>��"
            Response.Write "&nbsp;&nbsp;&nbsp;��<a href='javascript:history.go(-2)'>�ȵȣ�����ȷ��һ��</a>��" 
            Response.End         
                          
           Case "����" 
            Response.Write "<font size='4' color ='ff0000'>�����ڽ���һ���Σ�յĲ�������Ҫɾ�����ݣ���¼���Ϊ��ʷ���������Ҫ������</font><Br><Br>"
            Response.Write "��<a href='Process_del.asp?cat=2&ItmID="&request("ItmID")&"'>�������˵���������ī��ī�������ҵ���</a>��"
            Response.Write "&nbsp;&nbsp;&nbsp;��<a href='javascript:history.go(-2)'>�����ӣ��һ�����ȷ��һ�°�</a>��" 
            Response.End
            
          Case "GO"
             Response.Redirect "Process_Browse.asp?page="&Request("pagename")
          
          Case "����"                 
            SqlStr_c ="INSERT INTO t_Failbase_Connect(ItmID_Fb,Fb_His_ID,InUser,InTime) VALUES ('"
            SqlStr_c = SqlStr_c & request("ItmID") & "','" 
			SqlStr_c = SqlStr_c & request("Flb_His_ID") & "','"
			SqlStr_c = SqlStr_c & request.cookies("fmea.truename") & "','" 
			SqlStr_c = SqlStr_c & now() & "')"
            
            call openDB()
            conn.execute(SqlStr_c)
            call closeDB()

            Response.Redirect "Process_Browse.asp?page="&request("page")   
End Select

pgs =2500
select case request("cat")
       case "" 
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM v_Fmea_Item where PjID = "&request.cookies("fmea.PjID")&" order by SubPjName,PageNo,WorkQueue"
         pgs = 10
       
       '--����  
       case 1
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM v_Fmea_Item where (SubPjName like '%"&request("Sc")& _
           "%' Or PageNo like '%"&request("Sc")& _
		   "%' Or WorkTheme like '%"&request("Sc")& _
           "%' Or PartsNo like '%"&request("Sc")& _
           "%' Or PartsType like '%"&request("Sc")& _
           "%' Or PartsName like '%"&request("Sc")& _
           "%' Or WorkCata like '%"&request("Sc")& _
           "%' Or WorkContent like '%"&request("Sc")& _
           "%' Or NGMode like '%"&request("Sc")& _
           "%' Or NGEffect like '%"&request("Sc")& _
           "%' Or ProDesign like '%"&request("Sc")& _
           "%')  and PjID = '"&Request.cookies("fmea.PjID")&"' ORDER BY SubPjName,PageNo,WorkQueue" 
         
        '--���Ӧ��Ŀ  
        case 2
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM v_Fmea_Total where  RPN > 2  and ItmStatus = 'Y' and PjID = '"&Request.cookies("fmea.PjID")&"' ORDER BY SubPjName,PageNo,WorkQueue" 

       '--ȫϵ������ 
        case 3
         ttl = request.cookies("fmea.Serial")
         sql="SELECT * FROM v_Fmea_Total where (SubPjName like '%"&request("Sc")& _
           "%' Or PageNo like '%"&request("Sc")& _
		   "%' Or WorkTheme like '%"&request("Sc")& _
           "%' Or PartsNo like '%"&request("Sc")& _
           "%' Or PartsType like '%"&request("Sc")& _
           "%' Or PartsName like '%"&request("Sc")& _
           "%' Or WorkCata like '%"&request("Sc")& _
           "%' Or WorkContent like '%"&request("Sc")& _
           "%' Or NGMode like '%"&request("Sc")& _
           "%' Or NGEffect like '%"&request("Sc")& _
           "%' Or ProDesign like '%"&request("Sc")& _
		   "%') and  Serial = '"&Request.cookies("fmea.Serial")&"' ORDER BY SubPjName,PageNo,WorkQueue"  
        
        '--���Ʋ�Ʒ
        case 4
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM v_Fmea_Total where PartsOK is not null and PjID = '"&Request.cookies("fmea.PjID")&"' ORDER BY SubPjName,PageNo,WorkQueue"
     
        case 5
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM v_Fmea_Total where ItmID ="&request("ItmID")&" ORDER BY SubPjName,PageNo,WorkQueue" 
        
        '--RI��ѯ
        case 6
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM v_Fmea_Total where  RPN ="&request("Sc")&" and PjID = '"&Request.cookies("fmea.PjID")&"' ORDER BY SubPjName,PageNo,WorkQueue" 
      
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

var winconf = "menubar=no,toolbar=no,location=no,directories=no,status=no,scrollbar=no,resizeable=no,";

function about(filename,w,h) 
{
about = window.open(filename,"Copyright",winconf+"left=0,width="+w+",height="+h);
about.focus();
}
-->
</script> 
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
<form action="Process_Browse.asp">
<% 
pg = request("page")
if pg  ="" then  pg = 1
response.write "<input type='hidden' value='"& pg  &"' name='page'>"
%>
<div style ="float:right ;width��200px;hight:200px;">
<%
if request.cookies("fmea.power") > "1" then
  response.write "<input type='submit' value='��������' name='OK' class='smallInput'>&nbsp;"
  response.write "<input type='submit' value='����¼��' name='OK' class='smallInput'>&nbsp;"
  response.write "<input type='submit' value='��ʩ¼��' name='OK' class='smallInput'>&nbsp;"
end if
%>
<input name="Search" type="text" class="smallInput" size="40">
<input name="Vle" type="CheckBox" class="notline" Value ="Check">RIֵ
<input type="submit" value="����" name="OK" class="smallInput">
<input type="submit" value="ȫϵ������" name="OK" class="smallInput">
<input type="submit" value="ȫ��������" name="OK" class="smallInput">
<input type="submit" value="���Ʋ�Ʒ" name="OK" class="smallInput">
<input type="submit" value="���Ӧ��Ŀ" name="OK" class="smallInput">
<input type="submit" value="������" name="OK" class="smallInput">
&nbsp;&nbsp;
</div>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">

<tr><td width="75%" colspan=18 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">�ص���ҳ</a>]
[<a href="../PjManage/Pj_Browse.asp">����һ��</a>] 
[<a href="javascript:history.back()">����</a>] 
</td><td  width="50%" align="right">
 <% call clspage() %> 
</table>
</td></tr>

<tr bgcolor="#E6F7FF"><td height="30" align="center" colspan=18><b><% = ttl %>��ҵ���������б�</b></td></tr> 

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
<%
if request.cookies("fmea.power") > "1" then
   response.write "<td width='4%'>����</td>"
end if
%>
</tr>

<% 
if request("Act") = "Edit" then
    SqlStr_e = "SELECT * FROM t_Fmea_Item where ItmID ="&request("ItmID")
    Set rs_ed = Server.CreateObject("ADODB.Recordset")        
    rs_ed.open SqlStr_e,conn,1,1
    response.write"<tr bgcolor='#FCFCFC'>"
    response.write"<input name='ItmID' type='hidden' value ='"&rs_ed("ItmID")&"'>"
    response.write"<input name='oRPN' type='hidden' value ='"&rs_ed("RPN")&"'>"
    response.write"<td><input name='WorkQueue' class='notline' size='3' value ='"&rs_ed("WorkQueue")&"'></td>"
    response.write"<td><textarea name='SubPjName' cols='10' rows='4' wrap='virtual' class='notline' >"&rs_ed("SubPjName")&"</textarea></td>"
	response.write"<td><input name='PageNo' class='notline' size='9' value ='"&rs_ed("PageNo")&"'></td>" 
	response.write"<td>"
	
    response.write"<select name ='WorkTheme'>"
    response.write"<option value ='"&WorkTheme&"'>"&WorkTheme&"</option>"
    For i = 0 to UBound(ArWkThm) 
       response.write"<option value ='"&ArWkThm(i)&"'>"&ArWkThm(i)&"</option>"
    Next
    response.write"</select>"
	
	response.write"<td><input name='PartsNo' class='notline' size='15' value ='"&rs_ed("PartsNo")&"'></td>"    
    response.write"<td><div><select name='PartsType'>"
    response.write"<option value ='"&rs_ed("PartsType")&"'>"&rs_ed("PartsType")&"</option>"  
	For j = 0 to UBound(ArPt) 
       response.write"<option value ='"&ArPt(j)&"'>"&ArPt(j)&"</option>"
    Next
          
    response.write"</select></div><div>"
    response.write"<textarea name='PartsName' cols='15' rows='2' wrap='virtual' class='notline' >"&rs_ed("PartsName")&"</textarea></div></td>"
    response.write"<td><div><select name='WorkCata'>"
    response.write"<option value ='"&rs_ed("WorkCata")&"'>"&rs_ed("WorkCata")&"</option>"  
	For k = 0 to UBound(ArWkCt) 
       response.write"<option value ='"&ArWkCt(k)&"'>"&ArWkCt(k)&"</option>"
    Next
    response.write"</select></div>"
    response.write"<textarea name='WorkContent' cols='15' rows='2' wrap='virtual' class='notline' >"&rs_ed("WorkContent")&"</textarea></td>"
    response.write"<td><textarea name='NGMode' cols='22' rows='4' wrap='virtual' class='notline' >"&rs_ed("NGMode")&"</textarea></td>"
    response.write"<td><textarea name='NGEffect' cols='22' rows='4' wrap='virtual' class='notline' >"&rs_ed("NGEffect")&"</textarea></td>"
    response.write"<td><textarea name='ProDesign' cols='18' rows='4' wrap='virtual' class='notline' >"&rs_ed("ProDesign")&"</textarea></td>"
    response.write"<td><select name ='Sev' class= 'notline'>"
    response.write"<option value ='"&rs_ed("Sev")&"'>"&rs_ed("Sev")&"</option>"
    response.write"<option value ='1'>1</option>"
    response.write"<option value ='2'>2</option>"
    response.write"<option value ='3'>3</option>"
    response.write"<option value ='4'>4</option>"
    response.write"</select></td>"
    response.write"<td><select name ='Occ' class= 'notline'>"
    response.write"<option value ='"&rs_ed("Occ")&"'>"&rs_ed("Occ")&"</option>"
    response.write"<option value ='1'>1</option>"
    response.write"<option value ='2'>2</option>"
    response.write"<option value ='3'>3</option>"
    response.write"<option value ='4'>4</option>"
    response.write"</select></td>"

    response.write"<td><select name ='Det' class= 'notline'>"
    response.write"<option value ='"&rs_ed("Det")&"'>"&rs_ed("Det")&"</option>"
    response.write"<option value ='1'>1</option>"
    response.write"<option value ='2'>2</option>"
    response.write"<option value ='3'>3</option>"
    response.write"<option value ='4'>4</option>"
    response.write"</select></td>"

    response.write"<td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td></tr>"
    response.write"<tr bgcolor='#FCFCFC'><td colspan =16  align='center' height =35>"
    response.write"<input name='OK' value='����' type='submit' class='smallInput'>&nbsp;" 
    response.write"<input name='OK' value='�޶�(����)' type='submit' class='smallInput'>&nbsp;" 
    response.write"<input name='OK' value='��ֹ' type='submit' class='smallInput'>&nbsp;"
    response.write"<input name='OK' value='����' type='submit' class='smallInput'>&nbsp;"
    response.write"[<a href='javascript:history.back()'>����</a>]" 
    response.write"</td></tr>" 
    response.write"<tr bgcolor='#FCFCFC'><td colspan =16>"
    response.write""
    response.write"���������ǲ�������ԭ���ļ�¼�ģ����ԣ�ֻ��¼�����ʱʹ�ã�Ҫ���á�<br>"
    response.write"���޶�(����)���ǻ�����ԭ���ļ�¼�ģ��������޶���¼�ϲ鿴�����ԣ������ļ��޶�ʱʹ�á��޶���ɺ����ڡ������б��Ͻ������档<br>"
    response.write"����ֹ��������һ����������ԭ���ļ�¼�ģ���������ԭ�ļ��ϱ�����ǣ���ʱ���޶���ɺ�Ҳ��Ҫ�ڡ������б��Ͻ������档<br>"
    response.write"���������ǽ���¼ɾ���������������Ҫ��ʱ���ʹ�ã�����ȷ����ȷ�ϵġ�Ҫ���á�����á��������򡾷�ֹ����"
    response.write"</td></tr>" 
    call closeDB()
        
elseif request("Act") = "History" then
    SqlStr_e = "SELECT * FROM t_Fmea_Item where ItmID ="&request("ItmID")
    Set rs_ed = Server.CreateObject("ADODB.Recordset")        
    rs_ed.open SqlStr_e,conn,1,1
    response.write"<tr bgcolor='#FCFCFC'>"
    response.write"<input name='ItmID' type='hidden' value ='"&rs_ed("ItmID")&"'>"
    response.write"<td>"&rs_ed("WorkQueue")&"</td>"
    response.write"<td>"&rs_ed("SubPjName")&"</td>"
    response.write"<td>"&rs_ed("PageNo")&"</td>" 
    response.write"<td>"&rs_ed("WorkTheme")&"</td>" 		
    response.write"<td>"&rs_ed("PartsNo")&"</td>"    
    response.write"<td>("&rs_ed("PartsType")&")"&rs_ed("PartsName")&"</td>"
    response.write"<td>("&rs_ed("WorkCata")&")"&rs_ed("WorkContent")&"</td>"
    response.write"<td>"&rs_ed("NGMode")&"</td>"
    response.write"<td>"&rs_ed("NGEffect")&"</td>"
    response.write"<td>"&rs_ed("ProDesign")&"</td>"
    response.write"<td>"&rs_ed("Sev")&"</td>"
    response.write"<td>"&rs_ed("Occ")&"</td>"
    response.write"<td>"&rs_ed("Det")&"</td>"

    response.write"<td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td></tr>"
    response.write"<tr bgcolor='#FCFCFC'><td colspan =16  align='center' height =20>"
    response.write"���ڴ������ȥ����ı�ţ�<input name='Flb_His_ID'  class='smallInput'>&nbsp;" 
    response.write"</td></tr>" 
    response.write"<tr bgcolor='#FCFCFC'><td colspan =16  align='center' height =35>"
    response.write"<input name='OK' value='����' type='submit' class='smallInput'>&nbsp;" 
    response.write"[<a href='javascript:history.back()'>����</a>]" 
    response.write"</td></tr>" 
    response.write"<tr bgcolor='#FCFCFC'><td colspan =16>"
    response.write""
    response.write"�����롿��֮ǰ����Ĺ�ȥ�����ż��롣<br>"
    response.write"</td></tr>" 
    call closeDB()       
else

call MainBody() 

%>

</table>
</form>
<% call closeDB() %>
<% end if %>
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
       response.write"<div><a href=Advice_Browse.asp?cat=2&ItmID="&ItmID&">��Ӧ</a>&nbsp;"
       response.write"<a href= #  onClick=about('Advice_Read.asp?cat=2&ItmID="&ItmID&"',850,250)>�鿴</a></div>"
       if not isnull(rs("ItmID_his")) then response.write"<div><a href=../history/His_Process_Browse.asp?ItmID="&ItmID&">�޶�����</a></div>"
       if not isnull(rs("ItmID_fb"))  then response.write"<div><a href= # onClick=about('../His_Question/FBRead.asp?cat=3&ItmID="&ItmID&"',600,600)>��ȥ����</a></div>"
       response.write"</td>"       
    else
       response.write"<td>"& vRPN &"</td>" 
       response.write"<td>"
       if not isnull(rs("ItmID_his")) then response.write"<div><a href=../history/His_Process_Browse.asp?ItmID="&ItmID&">�޶�����</a></div>"
       if not isnull(rs("ItmID_fb")) then response.write"<div><a href= #  onClick=about('../His_Question/FBRead.asp?cat=3&ItmID="&ItmID&"',600,600)>��ȥ����</a></div>"
       response.write"</td>"
    end if 

    if request.cookies("fmea.power") > "1" then  
       response.write"<td>"
       response.write"<div height =20><a href=Process_Browse.asp?page="&page&"&Act=Edit&ItmID="&ItmID&">�޶�</a>&nbsp;"
       WQ = trim(rs("WorkQueue"))
       WQ = WQ + 0.01
       response.write"<a href=Process_Add.asp?cat=1&WQ="& WQ &">����</a></div>"
       if vRPN > 2 Then
         response.write"<div><a href= #  onClick=about('Advice_Add.asp?cat=1&ItmID="&ItmID&"&PjID="&rs("PjID")&"&PjKey="&rs("PjKey")&"',800,250)>��ʩ����</a></div>"
       end if
        response.write"<div height =20><a href=Process_Browse.asp?page="&page&"&Act=History&ItmID="&ItmID&">��ʷ����</a>&nbsp;"
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
	Response.Write "<a href=Process_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
	Response.Write "<a href=Process_Browse.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
	Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
	Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
	Response.Write "<a href=Process_Browse.asp?page=1>��ҳ</a>&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&pagecount&">βҳ</a>"
end if 

Response.Write "&nbsp;��<input type='Text' name='pagename' size='2' class='smallInput'  value ='"&page& "'>ҳ"
Response.Write "<input type='submit' value='GO' name='ok' class='smallInput'>"
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub
%>