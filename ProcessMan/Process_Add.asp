<!--#include file="../includes/db.asp"-->
<%
if  Request.cookies("fmea.PjID") = ""  then
    Response.Write "<script language='JavaScript' type='text/javascript'>alert('没有指定FMEA No.');</script>"    
    Response.Write "<a href ='..\PjManage\Pj_Browse.asp'>由于没有指定FMEA No.所以，请指定FMEA No.后录入相关内容</a>"
    Response.End
end if
   
Select Case Request("OK")  
       Case "回到列表"
           Response.redirect "Process_Browse.asp"

       Case "项目输入"

       if Request.cookies("fmea.PjID") = ""  then
           Response.Write "<a href ='..\PjManage\Pj_Browse.asp'>不好意思，内容出现丢失.请再指定FMEA No.后录入</a>"
           Response.End
       end if
       
       sql_ck = "select * from t_fmea_Item where PjID = '"&Request.cookies("fmea.PjID")&"'"
       sql_ck = sql_ck&" and  SubPjName ='"&request("SubPjName")&"'"
       sql_ck = sql_ck&" and PageNo='"&request("PageNo")&"'"
       sql_ck = sql_ck&" and WorkQueue='"&request("WorkQueue")&"'"
       
       call openDB()
       rs.open sql_ck,conn,1,1

       If rs.eof then
         Sev = int(request("Sev"))
         Occ = int(request("Occ"))
         Det = int(request("Det"))
         RPN = round((Sev * Occ * Det)^(1/3),1)
     
         ProDesign =request("Pro1")&request("Pro2")&request("Pro3")&request("Pro4")&request("ProDesign")

         sql= "Insert into t_Fmea_Item"
         sql= sql & "(PjID,Pjkey,SubPjName,PageNo,WorkQueue,WorkTheme,PartsNo,PartsName,PartsType,WorkCata,WorkContent,"
         sql= sql & "NGMode,NGEffect,ProDesign,Sev,Occ,Det,RPN,ItmStatus,InUser,InTime) values ('" 
         sql= sql & Request.cookies("fmea.PjID")&"','" 
         sql= sql & Request.cookies("fmea.Pjkey")&"','" 
         sql= sql & request("SubPjName")&"','" 
         sql= sql & request("PageNo")&"','" 
         sql= sql & request("WorkQueue")&"','"
		 sql= sql & request("WorkTheme")&"','" 
         sql= sql & request("PartsNo")&"','" 
         sql= sql & request("PartsName")&"','" 
         sql= sql & request("PartsType")&"','" 
         sql= sql & request("WorkCata")&"','" 
         sql= sql & request("WorkContent")&"','"  
         sql= sql & request("NGMode")&"','"
         sql= sql & request("NGEffect")&"','"
         sql= sql & ProDesign&"','"
         sql= sql & request("Sev")&"','"
         sql= sql & request("Occ")&"','"
         sql= sql & request("Det")&"','"
         sql= sql & RPN &"','Y','"
         sql= sql & request.cookies("fmea.truename") &"','"
         sql= sql & Now() &"')" 
    
         conn.execute(sql)
         call closeDB()
        
       Else
         response.write "<script language='JavaScript'>alert('输入页码及作业顺序重复!');</script>"
         Response.Write "<script language='javascript' type='text/javascript'>history.go(-1)</script>"
         response.end
       End if 
				       
End Select
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript" type="text/javascript">
<!--
function jump() {
 if(event.keyCode==13)event.keyCode=9
}

function showCustomer(str){
var xmlhttp;
if (str=="")
 {
 document.getElementById("txtHint").innerHTML="";
 return;
 }
if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
 xmlhttp=new XMLHttpRequest();
 }
else {// code for IE6, IE5
 xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
 }
xmlhttp.onreadystatechange=function() {
 if (xmlhttp.readyState==4 && xmlhttp.status==200) {
  document.getElementById("txtHint").innerHTML=xmlhttp.responseText;
    }
 }
xmlhttp.open("GET","getcustomer.html?q="+str,true);
xmlhttp.send();
}

-->
</script>
<title>项目输入</title>
</head>

<body marginheight="0" topmargin="10" bottommargin="0" leftmargin="0" bgcolor="#f9f9f9">
<form action="Process_Add.asp" method="post">
<table width="750"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center">
<tr bgcolor="#6c9c6c">
<td colspan ="2" height ="30" align ="center">
<font size ="3" color ="ffffff"><% = Request.cookies("fmea.PjKey") %>评价内容输入</font>
</td></tr>
<tr bgcolor="#FCFCFC">
<td width = "150">工程内容</td>
<td><input name="SubPjName" class="notline" size="40" value ="<% = Request("SubPjName") %>" tabindex ="1" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCFCFC">
<td width = "150">作业标准书页码</td>
<td><input name="PageNo" class="notline" size="25" value ="<% = Request("PageNo") %>" tabindex ="2" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCFCFC">
<td>作业顺序</td>
<td><input name="WorkQueue" class="notline" size="8" value ="<% = Request("WQ") %>"  tabindex ="3" onKeyDown="jump()">
</td></tr>
<tr bgcolor="#FCFCFC">
<td>作业场景</td>
<td>
<% 
  response.write"<select name ='WorkTheme'>"
  For i = 0 to UBound(ArWkThm) 
     response.write"<option value ='"&ArWkThm(i)&"'>"&ArWkThm(i)&"</option>"
  Next
  response.write"</select>"
%>  
<!-- <input name="WorkTheme" class="notline" size="50"  tabindex ="4" onKeyDown="jump()"> -->

</td></tr>
<tr bgcolor="#FCFCFC">
<td>部品编号</td>
<td><input name="PartsNo" class="notline" size="50"  tabindex ="5" onKeyDown="jump()"></td></tr>
<tr bgcolor="#FCFCFC">
<td>部品名称</td>
<td><input name="PartsName" class="notline" size="70"  tabindex ="6" onKeyDown="jump()"></td></tr>
<td bgcolor="#FCFCFC">部品种类</td>
<td bgcolor="#FCFCFC">
<div>
<input name="PartsType" type ="Radio" value ="框架件"  tabindex ="7" onKeyDown="jump()"> 框架件
<input name="PartsType" type ="Radio" value ="驱动件"> 驱动件
<input name="PartsType" type ="Radio" value ="光学件"> 光学件
<input name="PartsType" type ="Radio" value ="电气件"> 电气件
<input name="PartsType" type ="Radio" value ="引导件"> 引导件
<input name="PartsType" type ="Radio" value ="导通件"> 导通件
<input name="PartsType" type ="Radio" value ="组立部件"> 组立部件
</div><div>
<input name="PartsType" type ="Radio" value ="弹簧"> 弹簧
<input name="PartsType" type ="Radio" value ="轴承"> 轴承
<input name="PartsType" type ="Radio" value ="海绵"> 海绵
<input name="PartsType" type ="Radio" value ="胶片"> 胶片
<input name="PartsType" type ="Radio" value ="卡环"> 卡环
</div><div>
<input name="PartsType" type ="Radio" value ="标签"> 标签
<input name="PartsType" type ="Radio" value ="螺丝"> 螺丝
<input name="PartsType" type ="Radio" value ="同捆品"> 同捆品
<input name="PartsType" type ="Radio" value ="润滑材"> 润滑材
<input name="PartsType" type ="Radio" value ="耗材"> 耗材
<input name="PartsType" type ="Radio" value ="其他"> 其他
<input name="PartsType" type ="Radio" value ="-"> -
</div>
</td></tr>
<tr bgcolor="#FCFCFC">
<td>作业分类</td>
<td>
<select name="WorkCata" class="notline"  tabindex ="8" onKeyDown="jump()">
 <option value="打螺丝">打螺丝</option>
 <option value="装入">装入</option>
 <option value="张贴">张贴</option>
 <option value="结线">结线</option>
 <option value="线处理">线处理</option> 
 <option value="涂布">涂布</option>
 <option value="清扫">清扫</option>
 <option value="设定">设定</option>
 <option value="检查">检查</option>
 <option value="同捆">同捆</option>
 <option value="-">-</option>
</select>
</td></tr>
<tr bgcolor="#FCFCFC">
<td>作业内容</td>
<td><input name="WorkContent" class="notline" size="100"  tabindex ="9" onKeyDown="jump()"></td>
</tr>
<tr bgcolor="#FCFCFC">
<td>潜在缺陷模式</td>
<td><input name="NGMode" class="notline" size="100"  tabindex ="10" onKeyDown="jump()"></td></tr>
<tr bgcolor="#FCFCFC">
<td>潜在缺陷后果</td>
<td><input name="NGEffect" class="notline" size="100"  tabindex ="11" onKeyDown="jump()"></td></tr>
<tr bgcolor="#FCFCFC">
<td>现行工程设计及过程控制</td>
<td>
<input name="Pro1" type ="checkbox" value ="工程管理；">工程管理
<input name="Pro2" type ="checkbox" value ="顺次管理；">顺次管理
<input name="Pro3" type ="checkbox" value ="套数管理；">套数管理
<input name="Pro4" type ="checkbox" value ="数量管理；">数量管理
&nbsp;&nbsp;其他：<input name="ProDesign" class="notline" size="30" value ="目视检查" tabindex ="12" onKeyDown="jump()">

</td></tr>
<tr bgcolor="#FCFCFC">
<td>严重度Sev</td>
<td>
<input name="Sev" type ="Radio" value ="1"  tabindex ="13" onKeyDown="jump()"> 1
<input name="Sev" type ="Radio" value ="2"> 2
<input name="Sev" type ="Radio" value ="3"> 3
<input name="Sev" type ="Radio" value ="4"> 4

</td></tr>
<tr bgcolor="#FCFCFC">
<td>频度Occ</td>
<td>
<input name="Occ" type ="Radio" value ="1"  tabindex ="14" onKeyDown="jump()"> 1
<input name="Occ" type ="Radio" value ="2"> 2
<input name="Occ" type ="Radio" value ="3"> 3
<input name="Occ" type ="Radio" value ="4"> 4

</td></tr>
<tr bgcolor="#FCFCFC">
<td>不易探测度Det</td>
<td>
<input name="Det" type ="Radio" value ="1"  tabindex ="15" onKeyDown="jump()"> 1
<input name="Det" type ="Radio" value ="2"> 2
<input name="Det" type ="Radio" value ="3"> 3
<input name="Det" type ="Radio" value ="4"> 4

</td>
</tr>
<tr bgcolor="#FCFCFC" height ="30">
<td colspan = "2" align ="center" >
<input name="OK" value="项目输入" type="submit" class="smallInput"  tabindex ="16">
<input name="OK" value="回到列表" type="submit" class="smallInput"  tabindex ="17">
</td></tr>
</table>
</form>
</body>
</html>

