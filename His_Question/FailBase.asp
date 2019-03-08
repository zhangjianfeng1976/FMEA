<!--#include file="../includes/db.asp"-->
<% 
Select case request("OK")
    case "查询" 	                
          Select case Request("cata")
            case "模糊查询"
                Response.Redirect "FailBase.asp?Cat=0&Sc="&Request("Search")								
          End select         
    case "输入" 
        Response.Redirect "FailBase.asp?Action=ADD"
    case "GO"
        Response.Redirect "FailBase.asp?page="&Request("pagename")	
    case "确认"
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

     case "修订"		 
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

   
'删除______________________________________________________________________'
ID=request("ID")
Select Case Request("Action")
    Case "Delete" 
        Call openDB()
        sql="SELECT InUser FROM  t_FailBase where ID ="&ID
        rs.open sql,conn,1,1
	if trim(rs("InUser"))<>trim(Request.cookies("fmea.truename")) then
	   response.Write "你不能删除别人的记录!![<a href='javascript:history.back()'>返回</a>]"
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
	    response.Write "你不能修订别人的记录!![<a href='javascript:history.back()'>返回</a>]"
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
//双击鼠标滚动屏幕的代码
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
	font-family: "幼圆";
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
<option>模糊查询</option>
</select>
<input name="Search" type="text" class="smallInput" size="18">
<input type="submit" value="查询" name="OK" class="smallInput">
<% if request.cookies("fmea.power") >  "1" then 
response.write"<input type='submit' value='输入' name='OK' class='smallInput'>"
end if
%>
[<a href="javascript:history.back()">返回</a>]
</div><div style ="float:left;"><font face="华文中宋" size="5" color="#000000">过去问题</font></div>
</div>  
</td>

<tr Height="20"><td align="left"  bgcolor="#E6F7FF">
<div style ="float:left;">
<div style ="float:right ;">
<% call clspage() %>
</div><div style ="float:left;">
[<a href="../desktop.asp">回到首页</a>] [<a href="javascript:history.back()">返回</a>] 
</div></div>
</td></tr>

</tr><tr><td bgcolor='#fcfcfc'>

<% 
select case request("Action")
case "ADD"
  response.write"&nbsp;"
  response.write"<table width='60%' border='0' cellpadding='0' cellspacing='1' bgcolor='#5A8BCE' align='center'>"
  response.write"<tr bgcolor='#fcfcfc'>"
  response.write"<td width='20%'>机种</td><td><input name='Model' class='notline' size='15' tabindex ='1'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>"  
  response.write"<td width='20%'>日期</td><td><input name='Dtae' class='notline' size='15' tabindex ='2' value ='"&date()&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>制造</td>"
  response.write"<td>" 
  response.write"<input name='Mdepartment' type ='Radio' value ='机器1'>机器1&nbsp;"
  response.write"<input name='Mdepartment' type ='Radio' value ='机器2'>机器2&nbsp;"
  response.write"</td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>发生场所</td>"
  response.write"<td>" 
  response.write"<input name='LocationCata' type ='Radio' value ='LQC检出不良'>LQC检出不良&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='PQC不良'>PQC不良&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='GSA不良'>GSA不良&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='SET UP不良'>SET UP不良&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='轻微投诉'>轻微投诉&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='市场投诉'>市场投诉&nbsp;"
  response.write"<input name='LocationCata' type ='Radio' value ='其他'>其他&nbsp;"
  response.write"</td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>不良现象</td>"
  response.write"<td><font size='2' color='#0000FF'>"
  response.write"请规范填写格式：某某位置的某某部品发生何种不良导致何种现象 <br>"
  response.write"例：MAIN DRIVE UNIT上的CUT WASHER 4半插松脱导致发生异常音"
  response.write"</font>"
  response.write"<textarea name='FailContent' cols='100' rows='3' wrap='virtual' class='notline' ></textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>作业场景</td>"
  response.write"<td>"
  response.write"<select name ='WorkTheme'>"
  For i = 0 to UBound(ArWkThm) 
     response.write"<option value ='"&ArWkThm(i)&"'>"&ArWkThm(i)&"</option>"
  Next
  response.write"</select>"
  response.write"</td>" 
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>部品分类</td>"
  response.write"<td>"
  response.write"<div><input name='PartsType' type ='Radio' value ='框架件'>框架件&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='驱动件'>驱动件&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='光学件'>光学件&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='电气件'>电气件&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='引导件'>引导件&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='导通件'>导通件&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='组立部件'>组立部件&nbsp;"
  response.write"</div><div>"
  response.write"<input name='PartsType' type ='Radio' value ='弹簧'>弹簧&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='轴承'>轴承&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='海绵'>海绵&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='胶片'>胶片&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='卡环'>卡环&nbsp;"
  response.write"</div><div>"
  response.write"<input name='PartsType' type ='Radio' value ='标签'>标签&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='螺丝'>螺丝&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='同捆品'>同捆品&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='润滑材'>润滑材&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='耗材'>耗材&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='其他'>其他&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='-'>-&nbsp;"
  response.write"</div>"
  response.write"</td>"  
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>作业分类</td>"
  response.write"<td>"
  response.write"<div><input name='WorkCata' type ='Radio' value ='打螺丝'>打螺丝&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='装入'>装入&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='张贴'>张贴&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='结线'>结线&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='线处理'>线处理&nbsp;"
  response.write"</div><div>"
  response.write"<input name='WorkCata' type ='Radio' value ='涂布'>涂布&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='清扫'>清扫&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='设定'>设定&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='检查'>检查&nbsp;"
  response.write"<input name='WorkCata' type ='Radio' value ='同捆'>同捆&nbsp;"
  response.write"<input name='PartsType' type ='Radio' value ='-'>-&nbsp;"
  response.write"</div>"
  response.write"</td>" 
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>不良类别</td>"
  response.write"<td>" 
  response.write"<div><input name='FailCatalog' type ='Radio' value ='欠品'>欠品&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='漏涂油'>漏涂油&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='漏作业'>漏作业&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='仕样错误'>仕样错误&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='安装不到位'>安装不到位&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='方向错误'>方向错误&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='位置错误'>位置错误&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='结线半插斜插'>结线半插斜插&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='漏结线'>漏结线&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='夹线'>夹线&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='结线插歪PIN'>结线插歪PIN&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='线处理错误'>线处理错误&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='结线错误'>结线错误&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='粘贴超基准'>粘贴超基准&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='粘贴浮起'>粘贴浮起&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='螺丝半打斜打'>螺丝半打斜打&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='作业不正确'>作业不正确&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='设置错误'>设置错误&nbsp;"
  response.write"</div><div>"
  response.write"<input name='FailCatalog' type ='Radio' value ='部品不良'>部品不良&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='异品/异物混入'>异品/异物混入&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='涂油污染'>涂油污染&nbsp;"
  response.write"<input name='FailCatalog' type ='Radio' value ='外观不良'>外观不良&nbsp;"
  response.write"</div><div>"
  response.write"</td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>原因</td><td><textarea name='FailReason' cols='100' rows='6' wrap='virtual' class='notline' ></textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>对策</td><td><textarea name='FailAction' cols='100' rows='6' wrap='virtual' class='notline' ></textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>对应方法</td>"
  response.write"<td>"
  response.write"<div><input name='ActionMode1' type ='CheckBox' value ='工程管理及顺次管理；'>工程管理及顺次管理</div>"
  response.write"<div><input name='ActionMode2' type ='CheckBox' value ='配KIT；'>配KIT&nbsp;&nbsp;"
  response.write"<input name='ActionMode3' type ='CheckBox' value ='套数管理；'>套数管理&nbsp;&nbsp;"
  response.write"<input name='ActionMode4' type ='CheckBox' value ='数量管理；'>数量管理&nbsp;&nbsp;"
  response.write"<input name='ActionMode5' type ='CheckBox' value ='前配后装；'>前配后装&nbsp;&nbsp;"
  response.write"<input name='ActionMode51' type ='CheckBox' value ='称重管理；'>称重管理</div>"
  response.write"<div><input name='ActionMode6' type ='CheckBox' value ='MARKING；'>MARKING&nbsp;&nbsp;"
  response.write"<input name='ActionMode7' type ='CheckBox' value ='工程分割；'>工程分割&nbsp;&nbsp;"
  response.write"<input name='ActionMode8' type ='CheckBox' value ='OHP确认；'>OHP确认&nbsp;&nbsp;"
  response.write"<input name='ActionMode9' type ='CheckBox' value ='相似部品揭示牌；'>相似部品揭示牌</div>"
  response.write"<div><input name='ActionMode10' type ='CheckBox' value ='检查手顺书追加；'>检查手顺书追加&nbsp;&nbsp;"
  response.write"<input name='ActionMode11' type ='CheckBox' value ='作业指导书追加；'>作业指导书追加&nbsp;&nbsp;"
  response.write"<input name='ActionMode12' type ='CheckBox' value ='作业认定书追加；'>作业认定书追加</div>"
  response.write"<div><input name='ActionMode13' type ='CheckBox' value ='工具治具追加；'>工具治具追加</div>"
  response.write"</td></tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>FMEA检讨分类</td>"
  response.write"<td>" 
  response.write"<input name='FMEARsCata' type ='Radio' value ='制造对应遗漏'>制造对应遗漏&nbsp;"
  response.write"<input name='FMEARsCata' type ='Radio' value ='制造对应方法错误'>制造对应方法错误&nbsp;"
  response.write"<input name='FMEARsCata' type ='Radio' value ='用错作业场景'>用错作业场景&nbsp;"
  response.write"<input name='FMEARsCata' type ='Radio' value ='指南检讨遗漏'>指南检讨遗漏&nbsp;"
  response.write"<input name='FMEARsCata' type ='Radio' value ='其他'>其他&nbsp;"
  response.write"</td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>FMEA检讨内容</td><td><textarea name='FMEAResearch' cols='100' rows='4' wrap='virtual' class='notline'></textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td colspan ='2' Height ='40' align = 'center'><input type='submit' value='确认' name='OK' class='smallInput' tabindex ='9'></td>"
  response.write"</tr>"  
  response.write"</table>"
  response.write"&nbsp;"
case "Edit" 
  response.write"&nbsp;"
  response.write"<table width='60%' border='0' cellpadding='0' cellspacing='1' bgcolor='#5A8BCE' align='center'>"
  response.write"<input name='ID' type='hidden'  value = '"&ID&"'>"
  response.write"<tr bgcolor='#fcfcfc'>"
  response.write"<td>机种</td><td><input name='Model' class='notline' size='15' tabindex ='1' value = '"&Model&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>"  
  response.write"<td>日期</td><td><input name='Dtae' class='notline' size='15' tabindex ='2' value = '"&Dtae&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>制造</td><td>"
  response.write"<select name ='Mdepartment'>"
  response.write"<option value ='"&Mdepartment&"'>"&Mdepartment&"</option>"
  response.write"<option value ='机器1'>机器1</option>"
  response.write"<option value ='机器2'>机器2</option>"
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>发生场所</td><td>"
  response.write"<select name ='LocationCata'>"
  response.write"<option value ='"&LocationCata&"'>"&LocationCata&"</option>"
  response.write"<option value ='LQC检出不良'>LQC检出不良</option>"
  response.write"<option value ='PQC不良'>PQC不良</option>"
  response.write"<option value ='GSA不良'>GSA不良</option>"
  response.write"<option value ='SET UP不良'>SET UP不良</option>"
  response.write"<option value ='轻微投诉'>轻微投诉</option>"
  response.write"<option value ='市场投诉'>市场投诉</option>"
  response.write"<option value ='其他'>其他</option>"
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>不良现象</td><td><textarea name='FailContent' cols='100' rows='3' wrap='virtual' class='notline' >"&FailContent&"</textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>作业场景</td>"
  response.write"<td>" 
  response.write"<select name ='WorkTheme'>"
  response.write"<option value ='"&WorkTheme&"'>"&LWorkTheme&"</option>"
  For i = 0 to UBound(ArWkThm) 
     response.write"<option value ='"&ArWkThm(i)&"'>"&ArWkThm(i)&"</option>"
  Next
  response.write"</select>"
  response.write"</td>"  
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>部品分类</td><td>"
  response.write"<select name ='PartsType'>"
  response.write"<option value ='"&PartsType&"'>"&PartsType&"</option>"
  For j = 0 to UBound(ArPt) 
       response.write"<option value ='"&ArPt(j)&"'>"&ArPt(j)&"</option>"
  Next
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td width='20%'>作业分类</td><td>" 
  response.write"<select name ='WorkCata'>"
  response.write"<option value ='"&WorkCata&"'>"&WorkCata&"</option>"
  For k = 0 to UBound(ArWkCt) 
       response.write"<option value ='"&ArWkCt(k)&"'>"&ArWkCt(k)&"</option>"
  Next
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>不良类别</td><td>"
  response.write"<select name ='FailCatalog'>"
  response.write"<option value ='"&FailCatalog&"'>"&FailCatalog&"</option>"
  response.write"<option value ='欠品'>欠品</option>"
  response.write"<option value ='漏涂油'>漏涂油</option>"
  response.write"<option value ='漏作业'>漏作业</option>"
  response.write"<option value ='仕样错误'>仕样错误</option>"
  response.write"<option value ='安装不到位'>安装不到位</option>"
  response.write"<option value ='方向错误'>方向错误</option>"
  response.write"<option value ='位置错误'>位置错误</option>"
  response.write"<option value ='结线半插斜插'>结线半插斜插</option>"
  response.write"<option value ='漏结线'>漏结线</option>"
  response.write"<option value ='夹线'>夹线</option>"
  response.write"<option value ='结线插歪PIN'>结线插歪PIN</option>"
  response.write"<option value ='线处理错误'>线处理错误</option>"
  response.write"<option value ='结线错误'>结线错误</option>"
  response.write"<option value ='粘贴超基准'>粘贴超基准</option>"
  response.write"<option value ='粘贴浮起'>粘贴浮起</option>"
  response.write"<option value ='螺丝半打斜打'>螺丝半打斜打</option>"
  response.write"<option value ='作业不正确'>作业不正确</option>"
  response.write"<option value ='设置错误'>设置错误</option>"
  response.write"<option value ='部品不良'>部品不良</option>"
  response.write"<option value ='异品/异物混入'>异品/异物混入</option>"
  response.write"<option value ='涂油污染'>涂油污染</option>"
  response.write"<option value ='外观不良'>外观不良</option>"
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>原因</td><td><textarea name='FailReason' cols='100' rows='6' wrap='virtual' class='notline' >"&FailReason&"</textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>对策</td><td><textarea name='FailAction' cols='100' rows='6' wrap='virtual' class='notline' >"&FailAction&"</textarea></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>对应方法</td><td><input name='ActionMode' class='notline' size='100' tabindex ='8' value = '"&ActionMode&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>"  
  response.write"<td>FMEA检讨分类</td><td>"
  response.write"<select name ='FMEARsCata'>"
  response.write"<option value ='"&FMEARsCata&"'>"&FMEARsCata&"</option>"
  response.write"<option value ='制造对应遗漏'>制造对应遗漏</option>"
  response.write"<option value ='制造对应方法错误'>制造对应方法错误</option>"
  response.write"<option value ='用错作业场景'>用错作业场景</option>"
  response.write"<option value ='指南检讨遗漏'>指南检讨遗漏</option>"
  response.write"<option value ='其他'>其他</option>"
  response.write"</select></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td>FMEA检讨内容</td><td><input name='FMEAResearch' class='notline' size='100' tabindex ='6' value = '"&FMEAResearch&"'></td>"
  response.write"</tr><tr bgcolor='#fcfcfc'>" 
  response.write"<td colspan ='2' Height ='40' align ='center'>"
  response.write"<input type='submit' value='修订' name='OK' class='smallInput' tabindex ='9'></td>"
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
response.write "<td width='2%' height='20'>编号</td>"
response.write "<td width='2%' height='5'>内容</td>"
response.write "<td width='5%' height='20'>机种</td>"
response.write "<td width='5%'>日期</td>"
response.write "<td width='5%'>发生场所</td>"
response.write "<td width='15%'>不良现象</td>"
response.write "<td width='3%'>作业场景</td>"
response.write "<td width='3%'>部品分类</td>"
response.write "<td width='3%'>作业分类</td>"
response.write "<td width='5%'>不良类别</td>"
response.write "<td width='8%'>对应方法</td>"
if request.cookies("fmea.power") > "1" then  response.write "<td width='5%'>操作</td>"
response.write "</tr>"

if rs.eof then
  response.write "<tr><td width=""75%"" colspan=12>"
  response.write "没有数据。"
  response.write "</td></tr>"
end if

n=1 
While Not rs.eof and n<=rs.pagesize 
  response.write"<tr bgcolor='#FCFCFC'>"
  'response.write"<td>"&rs("ID")&"</td>"
  response.write"<td><a href=../ProcessMan/Process_Browse_total.asp?cat=6&ID="&rs("ID")&">"&rs("ID")&"</a></td>"
  response.write"<td><a href=FBRead.asp?cat=1&ID="&rs("ID")&">详细</a></td>"
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
	response.write"<td><a href=FailBase.asp?Action=Delete&id="&rs("id")&">删除</a>&nbsp;&nbsp;"
	response.write"<a href=FailBase.asp?Action=Edit&id="&rs("id")&">修订</a></td>"
  end if
  response.write"</tr>"
  n=n+1 
  rs.movenext '显示页面的数据' 
Wend 
response.write "</table>"

End Sub

Sub clspage()

if page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
	Response.Write "首页&nbsp;上页&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&page+1&">下页<a>&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&pagecount&">尾页<a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
	Response.Write "<a href=FailBase.asp?page=1>首页<a>&nbsp;&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&page-1&">上页<a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
	Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
	Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
	Response.Write "<a href=FailBase.asp?page=1>首页<a>&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&page-1&">上页<a>&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&page+1&">下页<a>&nbsp;"
	Response.Write "<a href=FailBase.asp?page="&pagecount&">尾页<a>"
end if

Response.Write "&nbsp;第<input type='Text' name='pagename' size='2' class='smallInput'>页"
Response.Write "<input type='submit' value='GO' name='ok' class='smallInput'>"
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"

end sub
%>

 
