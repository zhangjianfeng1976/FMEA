<!--#include file="../includes/db.asp"-->
<% 
Select case request("OK")
    case "��ѯ" 	                
          Select case Request("cata")
            case "ģ����ѯ"
                Response.Redirect "FailBase.asp?Cat=0&Sc="&Request("Search")								
          End select         
    case "����" 
        Response.Redirect "FailBase.asp?Action=ADD"
    case "GO"
        Response.Redirect "FailBase.asp?page="&Request("pagename")	
    case "ȷ��"
         ActM1   = request("ActionMode1")
         ActM2   = request("ActionMode2")
         ActM3   = request("ActionMode3")
         ActM4   = request("ActionMode4")
         ActM5   = request("ActionMode5")
         ActM51  = request("ActionMode51")
         ActM6   = request("ActionMode6")
         ActM7   = request("ActionMode7")
         ActM8   = request("ActionMode8")
         ActM9   = request("ActionMode9")
         ActM10  = request("ActionMode10")
         ActM11  = request("ActionMode11")
         ActM12  = request("ActionMode12")
         ActM13  = request("ActionMode13")
       
         ActMode =ActM1&ActM2&ActM3&ActM4&ActM5&ActM51&ActM6&ActM7&ActM8&ActM9&ActM10&ActM11&ActM12&ActM13
         sql="INSERT INTO t_FailBase(Model,Dtae,Mdepartment,LocationCata,WorkTheme,PartsType,WorkCata,FailContent,FailCatalog, " _
		     &"FailReason,FailAction,ActionMode,FMEARsCata,FMEAResearch,InUser,InTime) " _
             &"VALUES ('" _
         &request("Model")&"', '"_
	     &request("Dtae")&"', '"_
		 &request("Mdepartment")&"', '"_
	     &request("LocationCata")&"', '"_
		 &request("WorkTheme")&"', '"_
		 &request("PartsType")&"', '"_
		 &request("WorkCata")&"', '"_
	     &request("FailContent")&"', '"_
         &request("FailCatalog")&"', '"_
	     &request("FailReason")&"', '"_
	     &request("FailAction")&"', '"_
	     &ActMode&"', '"_
		 &request("FMEARsCata")&"', '"_
	     &request("FMEAResearch")&"', '"_
	     &Request.cookies("fmea.truename")&"', '"_
	     &Now()&"')"

	     call openDB()
	     conn.execute(sql)
	     call closeDB()
         response.Redirect "FailBase.asp"

     case "�޶�"		 
        sql="Update t_FailBase set Model = '"&request("Model") _
              &"',Dtae = '"&request("Dtae") _
			  &"',Mdepartment = '"&request("Mdepartment") _
              &"',LocationCata = '"&request("LocationCata") _
			  &"',WorkTheme = '"&request("WorkTheme") _
			  &"',PartsType = '"&request("PartsType") _
			  &"',WorkCata = '"&request("WorkCata") _
              &"',FailContent  = '"&request("FailContent") _
              &"',FailCatalog = '"&request("FailCatalog") _
              &"',FailReason = '"&request("FailReason") _
              &"',FailAction  = '"&request("FailAction") _
              &"',ActionMode  = '"&request("ActionMode") _
			  &"',FMEARsCata  = '"&request("FMEARsCata") _
              &"',FMEAResearch  = '"&request("FMEAResearch") _
              &"',InUser = '"&Request.cookies("fmea.truename") _
              &"',InTime = '"&Now()&"' where Id = "&request("ID")
        call openDB()
        conn.execute(sql)
        call closeDB()
        response.Redirect "FailBase.asp"		 		  	  
 End Select

   
'ɾ��______________________________________________________________________'
ID=request("ID")
Select Case Request("Action")
    Case "Delete" 
        Call openDB()
        sql="SELECT InUser FROM  t_FailBase where ID ="&ID
        rs.open sql,conn,1,1
	if trim(rs("InUser"))<>trim(Request.cookies("fmea.truename")) then
	   response.Write "�㲻��ɾ�����˵ļ�¼!![<a href='javascript:history.back()'>����</a>]"
	   response.End()
	   Call closeDB()
	else
	   Call closeDB()
           sql="Delete t_FailBase Where ID ="&ID       
           Call OpenDB()
           conn.execute(sql)
           Call closeDB()
	end if
	   
    Case "Edit" 
        Call openDB()
	sql="SELECT  * FROM  t_FailBase where ID ="&ID
        rs.open sql,conn,1,1
	if trim(rs("InUser")) <> trim(Request.cookies("fmea.truename")) then
	    response.Write "�㲻���޶����˵ļ�¼!![<a href='javascript:history.back()'>����</a>]"
	    response.End()
	    Call closeDB()
	else
	     ID             = rs("ID")
         Model          = rs("Model")
         Dtae           = rs("Dtae")
		 Mdepartment    = rs("Mdepartment")
         LocationCata   = rs("LocationCata")
		 WorkTheme      = rs("WorkTheme")
		 PartsType      = rs("PartsType")
		 WorkCata       = rs("WorkCata")
         FailContent    = rs("FailContent")
         FailCatalog    = rs("FailCatalog")
         FailReason     = rs("FailReason")
         FailAction     = rs("FailAction")
         ActionMode     = rs("ActionMode")
		 FMEARsCata     = rs("FMEARsCata")
         FMEAResearch   = rs("FMEAResearch")
         InUser         = rs("InUser")
         InTime         = rs("InTime")
         call closeDB()    
	end if	   
End Select

pgx=5000
select case request("cat")
    case ""
        pgx=25
        sql="SELECT * FROM t_FailBase order by id desc"
    case 0
        sql="SELECT * FROM t_FailBase where Model like '%"&request("Sc")& _
            "%' Or LocationCata  like '%"&request("Sc")& _
			"%' Or WorkTheme     like '%"&request("Sc")& _
			"%' Or PartsType     like '%"&request("Sc")& _
			"%' Or WorkCata      like '%"&request("Sc")& _
            "%' Or FailContent   like '%"&request("Sc")& _
            "%' Or FailCatalog   like '%"&request("Sc")& _
            "%' Or FailReason    like '%"&request("Sc")& _
            "%' Or FailAction    like '%"&request("Sc")& _
            "%' Or ActionMode    like '%"&request("Sc")& _
			"%' Or FMEARsCata    like '%"&request("Sc")& _
            "%' Or FMEAResearch  like '%"&request("Sc")& _
	        "%' Or InUser like '%"&request("Sc")& _
            "%' ORDER BY dtae DESC"
    case 1
        sql="SELECT * FROM t_FailBase where LocationCata like '%"&request("Sc")&"%' order by dtae desc"
    case 2
        sql="SELECT * FROM t_FailBase where FailCatalog  like '%"&request("Sc")&"%' order by dtae desc"	
    case 3
        sql="SELECT * FROM t_FailBase where ItmID_fb = "&request("ItmID")	   				      
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Includes/Style.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript type=text/JavaScript>
//˫����������Ļ�Ĵ���
var currentpos,timer;
function initialize()
{
timer=setInterval ("scrollwindow ()",1);
}
function sc()
{
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos !=document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize
</SCRIPT>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 20;
	color: #FFFFFF;
	font-family: "��Բ";
}
-->
</style>
</head>

<body>
<form action="FailBase.asp" method = "post">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center" >
<tr Height ="30"><td align="right" bgcolor="#E6F7F0">
<div style ="float:left;">
<div style ="float:right ;">
<select name='cata' class="smallinput">
<option>ģ����ѯ</option>
</select>
<input name="Search" type="text" class="smallInput" size="18">
<input type="submit" value="��ѯ" name="OK" class="smallInput">
<% if request.cookies("fmea.power") >  "1" then 
response.write"<input type='submit' value='����' name='OK' class='smallInput'>"
end if
%>
[<a href="javascript:history.back()">����</a>]
</div><div style ="float:left;"><font face="��������" size="5" color="#000000">��ȥ����</font></div>
</div>  
</td>

<tr Height="20"><td align="left"  bgcolor="#E6F7FF">
<div style ="float:left;">
<div style ="float:right ;">
<% call clspage() %>
</div><div style ="float:left;">
[<a href="../desktop.asp">�ص���ҳ</a>] [<a href="javascript:history.back()">����</a>] 
</div></div>
</td></tr>

</tr><tr><td bgcolor='#fcfcfc'>

<% 
select case request("Action")
case "ADD"
  response.write"&nbsp;"
  response.write"<table width='60%' border='0' cellpadding='0' cellspacing='1' bgcolor='#5A8BCE' align='center'>"
  response.write"<tr bgcolor='#fcfcfc'>"
  response.write"<td width='20%'>����</td><td><input name='Model' class='notline' size='15' tabindex ='1'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>"  
  response.write"<td width='20%'>����</td><td><input name='Dtae' class='notline' size='15' tabindex ='2' value ='"&date()&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>����</td>"
  response.write"<td>" 
  response.write"<input name='Mdepartment' type ='Radio' value ='����1'>����1&nbsp;"
  response.write"<input name='Mdepartment' type ='Radio' value ='����2'>����2&nbsp;"
  response.write"</td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>��������</td>"
  response.write"<td>" 
  response.write"<input name='LocationCata' type ='Radio' value ='LQC�������'>LQC�������&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='PQC����'>PQC����&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='GSA����'>GSA����&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='SET UP����'>SET UP����&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='��΢Ͷ��'>��΢Ͷ��&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='�г�Ͷ��'>�г�Ͷ��&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='����'>����&nbsp;"
  response.write"</td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>��������</td>"
  response.write"<td><font size='2' color='#0000FF'>"
  response.write"��淶��д��ʽ��ĳĳλ�õ�ĳĳ��Ʒ�������ֲ������º������� <br>"
  response.write"����MAIN DRIVE UNIT�ϵ�CUT WASHER 4������ѵ��·����쳣��"
  response.write"</font>"
  response.write"<textarea name='FailContent' cols='100' rows='3' wrap='virtual' class='notline' ></textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>��ҵ����</td>"
  response.write"<td>"
  response.write"<select name ='WorkTheme'>"
  For i = 0 to UBound(ArWkThm) 
     response.write"<option value ='"&ArWkThm(i)&"'>"&ArWkThm(i)&"</option>"
  Next
  response.write"</select>"
  response.write"</td>" 
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>��Ʒ����</td>"
  response.write"<td>"
  response.write"<div><input name='PartsType' type ='Radio' value ='��ܼ�'>��ܼ�&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='������'>������&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='��ѧ��'>��ѧ��&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='������'>������&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='������'>������&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='��ͨ��'>��ͨ��&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='��������'>��������&nbsp;"
  response.write"</div><div>"
  response.write"<input name='PartsType' type ='Radio' value ='����'>����&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='���'>���&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='����'>����&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='��Ƭ'>��Ƭ&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='����'>����&nbsp;"
  response.write"</div><div>"
  response.write"<input name='PartsType' type ='Radio' value ='��ǩ'>��ǩ&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='��˿'>��˿&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='ͬ��Ʒ'>ͬ��Ʒ&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='�󻬲�'>�󻬲�&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='�Ĳ�'>�Ĳ�&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='����'>����&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='-'>-&nbsp;"
  response.write"</div>"
  response.write"</td>"  
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>��ҵ����</td>"
  response.write"<td>"
  response.write"<div><input name='WorkCata' type ='Radio' value ='����˿'>����˿&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='װ��'>װ��&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='����'>����&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='����'>����&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='�ߴ���'>�ߴ���&nbsp;"
  response.write"</div><div>"
  response.write"<input name='WorkCata' type ='Radio' value ='Ϳ��'>Ϳ��&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='��ɨ'>��ɨ&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='�趨'>�趨&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='���'>���&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='ͬ��'>ͬ��&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='-'>-&nbsp;"
  response.write"</div>"
  response.write"</td>" 
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>�������</td>"
  response.write"<td>" 
  response.write"<div><input name='FailCatalog' type ='Radio' value ='ǷƷ'>ǷƷ&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='©Ϳ��'>©Ϳ��&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='©��ҵ'>©��ҵ&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='��������'>��������&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='��װ����λ'>��װ����λ&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='�������'>�������&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='λ�ô���'>λ�ô���&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='���߰��б��'>���߰��б��&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='©����'>©����&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='����'>����&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='���߲���PIN'>���߲���PIN&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='�ߴ������'>�ߴ������&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='���ߴ���'>���ߴ���&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='ճ������׼'>ճ������׼&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='ճ������'>ճ������&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='��˿���б��'>��˿���б��&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='��ҵ����ȷ'>��ҵ����ȷ&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='���ô���'>���ô���&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='��Ʒ����'>��Ʒ����&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='��Ʒ/�������'>��Ʒ/�������&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='Ϳ����Ⱦ'>Ϳ����Ⱦ&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='��۲���'>��۲���&nbsp;"
  response.write"</div><div>"
  response.write"</td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>ԭ��</td><td><textarea name='FailReason' cols='100' rows='6' wrap='virtual' class='notline' ></textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>�Բ�</td><td><textarea name='FailAction' cols='100' rows='6' wrap='virtual' class='notline' ></textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>��Ӧ����</td>"
  response.write"<td>"
  response.write"<div><input name='ActionMode1' type ='CheckBox' value ='���̹���˳�ι���'>���̹���˳�ι���</div>"
  response.write"<div><input name='ActionMode2' type ='CheckBox' value ='��KIT��'>��KIT&nbsp;&nbsp;"
  response.write"<input name='ActionMode3' type ='CheckBox' value ='��������'>��������&nbsp;&nbsp;"
  response.write"<input name='ActionMode4' type ='CheckBox' value ='��������'>��������&nbsp;&nbsp;"
  response.write"<input name='ActionMode5' type ='CheckBox' value ='ǰ���װ��'>ǰ���װ&nbsp;&nbsp;"
  response.write"<input name='ActionMode51' type ='CheckBox' value ='���ع���'>���ع���</div>"
  response.write"<div><input name='ActionMode6' type ='CheckBox' value ='MARKING��'>MARKING&nbsp;&nbsp;"
  response.write"<input name='ActionMode7' type ='CheckBox' value ='���̷ָ'>���̷ָ�&nbsp;&nbsp;"
  response.write"<input name='ActionMode8' type ='CheckBox' value ='OHPȷ�ϣ�'>OHPȷ��&nbsp;&nbsp;"
  response.write"<input name='ActionMode9' type ='CheckBox' value ='���Ʋ�Ʒ��ʾ�ƣ�'>���Ʋ�Ʒ��ʾ��</div>"
  response.write"<div><input name='ActionMode10' type ='CheckBox' value ='�����˳��׷�ӣ�'>�����˳��׷��&nbsp;&nbsp;"
  response.write"<input name='ActionMode11' type ='CheckBox' value ='��ҵָ����׷�ӣ�'>��ҵָ����׷��&nbsp;&nbsp;"
  response.write"<input name='ActionMode12' type ='CheckBox' value ='��ҵ�϶���׷�ӣ�'>��ҵ�϶���׷��</div>"
  response.write"<div><input name='ActionMode13' type ='CheckBox' value ='�����ξ�׷�ӣ�'>�����ξ�׷��</div>"
  response.write"</td></tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>FMEA���ַ���</td>"
  response.write"<td>" 
  response.write"<input name='FMEARsCata' type ='Radio' value ='�����Ӧ��©'>�����Ӧ��©&nbsp;"
  response.write"<input name='FMEARsCata' type ='Radio' value ='�����Ӧ��������'>�����Ӧ��������&nbsp;"
  response.write"<input name='FMEARsCata' type ='Radio' value ='�ô���ҵ����'>�ô���ҵ����&nbsp;"
  response.write"<input name='FMEARsCata' type ='Radio' value ='ָ�ϼ�����©'>ָ�ϼ�����©&nbsp;"
  response.write"<input name='FMEARsCata' type ='Radio' value ='����'>����&nbsp;"
  response.write"</td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>FMEA��������</td><td><textarea name='FMEAResearch' cols='100' rows='4' wrap='virtual' class='notline'></textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td colspan ='2' Height ='40' align = 'center'><input type='submit' value='ȷ��' name='OK' class='smallInput' tabindex ='9'></td>"
  response.write"</tr>"  
  response.write"</table>"
  response.write"&nbsp;"
case "Edit" 
  response.write"&nbsp;"
  response.write"<table width='60%' border='0' cellpadding='0' cellspacing='1' bgcolor='#5A8BCE' align='center'>"
  response.write"<input name='ID' type='hidden'  value = '"&ID&"'>"
  response.write"<tr bgcolor='#fcfcfc'>"
  response.write"<td>����</td><td><input name='Model' class='notline' size='15' tabindex ='1' value = '"&Model&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>"  
  response.write"<td>����</td><td><input name='Dtae' class='notline' size='15' tabindex ='2' value = '"&Dtae&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>����</td><td>"
  response.write"<select name ='Mdepartment'>"
  response.write"<option value ='"&Mdepartment&"'>"&Mdepartment&"</option>"
  response.write"<option value ='����1'>����1</option>"
  response.write"<option value ='����2'>����2</option>"
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>��������</td><td>"
  response.write"<select name ='LocationCata'>"
  response.write"<option value ='"&LocationCata&"'>"&LocationCata&"</option>"
  response.write"<option value ='LQC�������'>LQC�������</option>"
  response.write"<option value ='PQC����'>PQC����</option>"
  response.write"<option value ='GSA����'>GSA����</option>"
  response.write"<option value ='SET UP����'>SET UP����</option>"
  response.write"<option value ='��΢Ͷ��'>��΢Ͷ��</option>"
  response.write"<option value ='�г�Ͷ��'>�г�Ͷ��</option>"
  response.write"<option value ='����'>����</option>"
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>��������</td><td><textarea name='FailContent' cols='100' rows='3' wrap='virtual' class='notline' >"&FailContent&"</textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>��ҵ����</td>"
  response.write"<td>" 
  response.write"<select name ='WorkTheme'>"
  response.write"<option value ='"&WorkTheme&"'>"&LWorkTheme&"</option>"
  For i = 0 to UBound(ArWkThm) 
     response.write"<option value ='"&ArWkThm(i)&"'>"&ArWkThm(i)&"</option>"
  Next
  response.write"</select>"
  response.write"</td>"  
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>��Ʒ����</td><td>"
  response.write"<select name ='PartsType'>"
  response.write"<option value ='"&PartsType&"'>"&PartsType&"</option>"
  For j = 0 to UBound(ArPt) 
       response.write"<option value ='"&ArPt(j)&"'>"&ArPt(j)&"</option>"
  Next
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>��ҵ����</td><td>" 
  response.write"<select name ='WorkCata'>"
  response.write"<option value ='"&WorkCata&"'>"&WorkCata&"</option>"
  For k = 0 to UBound(ArWkCt) 
       response.write"<option value ='"&ArWkCt(k)&"'>"&ArWkCt(k)&"</option>"
  Next
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>�������</td><td>"
  response.write"<select name ='FailCatalog'>"
  response.write"<option value ='"&FailCatalog&"'>"&FailCatalog&"</option>"
  response.write"<option value ='ǷƷ'>ǷƷ</option>"
  response.write"<option value ='©Ϳ��'>©Ϳ��</option>"
  response.write"<option value ='©��ҵ'>©��ҵ</option>"
  response.write"<option value ='��������'>��������</option>"
  response.write"<option value ='��װ����λ'>��װ����λ</option>"
  response.write"<option value ='�������'>�������</option>"
  response.write"<option value ='λ�ô���'>λ�ô���</option>"
  response.write"<option value ='���߰��б��'>���߰��б��</option>"
  response.write"<option value ='©����'>©����</option>"
  response.write"<option value ='����'>����</option>"
  response.write"<option value ='���߲���PIN'>���߲���PIN</option>"
  response.write"<option value ='�ߴ������'>�ߴ������</option>"
  response.write"<option value ='���ߴ���'>���ߴ���</option>"
  response.write"<option value ='ճ������׼'>ճ������׼</option>"
  response.write"<option value ='ճ������'>ճ������</option>"
  response.write"<option value ='��˿���б��'>��˿���б��</option>"
  response.write"<option value ='��ҵ����ȷ'>��ҵ����ȷ</option>"
  response.write"<option value ='���ô���'>���ô���</option>"
  response.write"<option value ='��Ʒ����'>��Ʒ����</option>"
  response.write"<option value ='��Ʒ/�������'>��Ʒ/�������</option>"
  response.write"<option value ='Ϳ����Ⱦ'>Ϳ����Ⱦ</option>"
  response.write"<option value ='��۲���'>��۲���</option>"
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>ԭ��</td><td><textarea name='FailReason' cols='100' rows='6' wrap='virtual' class='notline' >"&FailReason&"</textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>�Բ�</td><td><textarea name='FailAction' cols='100' rows='6' wrap='virtual' class='notline' >"&FailAction&"</textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>��Ӧ����</td><td><input name='ActionMode' class='notline' size='100' tabindex ='8' value = '"&ActionMode&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>"  
  response.write"<td>FMEA���ַ���</td><td>"
  response.write"<select name ='FMEARsCata'>"
  response.write"<option value ='"&FMEARsCata&"'>"&FMEARsCata&"</option>"
  response.write"<option value ='�����Ӧ��©'>�����Ӧ��©</option>"
  response.write"<option value ='�����Ӧ��������'>�����Ӧ��������</option>"
  response.write"<option value ='�ô���ҵ����'>�ô���ҵ����</option>"
  response.write"<option value ='ָ�ϼ�����©'>ָ�ϼ�����©</option>"
  response.write"<option value ='����'>����</option>"
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>FMEA��������</td><td><input name='FMEAResearch' class='notline' size='100' tabindex ='6' value = '"&FMEAResearch&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td colspan ='2' Height ='40' align ='center'>"
  response.write"<input type='submit' value='�޶�' name='OK' class='smallInput' tabindex ='9'></td>"
  response.write"</tr>"  
  response.write"</table>"
  response.write"&nbsp;"
case else  
  call mainbody
End Select 
 %>
</td></tr>
</table>
<% call closeDB() %>
</form>
</body>
</html>

<%
Sub Mainbody()

response.write"<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#5A8BCE' align='center'>"
response.write "<tr align='center' bgcolor='#E6F7FF'>"
response.write "<td width='2%' height='20'>���</td>"
response.write "<td width='2%' height='5'>����</td>"
response.write "<td width='5%' height='20'>����</td>"
response.write "<td width='5%'>����</td>"
response.write "<td width='5%'>��������</td>"
response.write "<td width='15%'>��������</td>"
response.write "<td width='3%'>��ҵ����</td>"
response.write "<td width='3%'>��Ʒ����</td>"
response.write "<td width='3%'>��ҵ����</td>"
response.write "<td width='5%'>�������</td>"
response.write "<td width='8%'>��Ӧ����</td>"
if request.cookies("fmea.power") > "1" then  response.write "<td width='5%'>����</td>"
response.write "</tr>"

if rs.eof then
  response.write "<tr><td width=""75%"" colspan=12>"
  response.write "û�����ݡ�"
  response.write "</td></tr>"
end if

n=1 
While Not rs.eof and n<=rs.pagesize 
  response.write"<tr bgcolor='#FCFCFC'>"
  'response.write"<td>"&rs("ID")&"</td>"
  response.write"<td><a href=../ProcessMan/Process_Browse_total.asp?cat=6&ID="&rs("ID")&">"&rs("ID")&"</a></td>"
  response.write"<td><a href=FBRead.asp?cat=1&ID="&rs("ID")&">��ϸ</a></td>"
  response.write"<td>"&rs("Model")&"</td>"
  response.write"<td>"&rs("Dtae")&"</td>"
  response.write"<td><a href=FailBase.asp?cat=1&sc="&rs("LocationCata")&">"&rs("LocationCata")&"</a></td>"
  response.write"<td>"&rs("FailContent")&"</td>"
  response.write"<td>"&rs("WorkTheme")&"</td>"
  response.write"<td>"&rs("PartsType")&"</td>"
  response.write"<td>"&rs("WorkCata")&"</td>"
  response.write"<td><a href=FailBase.asp?cat=2&sc="&rs("FailCatalog")&">"&rs("FailCatalog")&"</a></td>"
  response.write"<td>"&rs("ActionMode")&"</td>"
  if request.cookies("fmea.power") > "1" then  
	response.write"<td><a href=FailBase.asp?Action=Delete&id="&rs("id")&">ɾ��</a>&nbsp;&nbsp;"
	response.write"<a href=FailBase.asp?Action=Edit&id="&rs("id")&">�޶�</a></td>"
  end if
  response.write"</tr>"
  n=n+1 
  rs.movenext '��ʾҳ�������' 
Wend 
response.write "</table>"

End Sub

Sub clspage()

if page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
	Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&page+1&">��ҳ<a>&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&pagecount&">βҳ<a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
	Response.Write "<a href=FailBase.asp?page=1>��ҳ<a>&nbsp;&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&page-1&">��ҳ<a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
	Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
	Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
	Response.Write "<a href=FailBase.asp?page=1>��ҳ<a>&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&page-1&">��ҳ<a>&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&page+1&">��ҳ<a>&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&pagecount&">βҳ<a>"
end if

Response.Write "&nbsp;��<input type='Text' name='pagename' size='2' class='smallInput'>ҳ"
Response.Write "<input type='submit' value='GO' name='ok' class='smallInput'>"
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"

end sub
%>

 
