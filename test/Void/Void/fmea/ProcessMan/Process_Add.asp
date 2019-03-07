<!--#include file="../includes/db.asp"-->
<%
if  Request("PjID") = ""  then
    Response.Write "<script language='JavaScript' type='text/javascript'>alert('没有指定FMEA No.');</script>"    
    Response.Write "<a href ='..\PjManage\Pj_Browse.asp'>由于没有指定FMEA No.所以，请指定FMEA No.后录入相关内容</a>"
    Response.End
end if
   
Select Case Request("OK")  
       Case "项目输入"

       if Request("PjID") = ""  then
           Response.Write "<a href ='..\PjManage\Pj_Browse.asp'>不好意思，内容出现丢失.请再指定FMEA No.后录入</a>"
           Response.End
       end if
 
       Sev = int(request("Sev"))
       Occ = int(request("Occ"))
       Det = int(request("Det"))
       RPN = Sev * Occ * Det
       sql= "Insert into t_Fmea_Item"
       sql= sql & "(PjID,SubPjName,PageNo,WorkQueue,PartsNo,PartsName,PartsType,WorkCata,WorkContent,"
       sql= sql & "NGMode,NGEffect,ProDesign,Sev,Occ,Det,RPN,ItmStatus,InUser,InTime) values ('" 
       sql= sql & request("PjID")&"','" 
       sql= sql & request("SubPjName")&"','" 
       sql= sql & request("PageNo")&"','" 
       sql= sql & request("WorkQueue")&"','"
       sql= sql & request("PartsNo")&"','" 
       sql= sql & request("PartsName")&"','" 
       sql= sql & request("PartsType")&"','" 
       sql= sql & request("WorkCata")&"','" 
       sql= sql & request("WorkContent")&"','"  
       sql= sql & request("NGMode")&"','"
       sql= sql & request("NGEffect")&"','"
       sql= sql & request("ProDesign")&"','"
       sql= sql & request("Sev")&"','"
       sql= sql & request("Occ")&"','"
       sql= sql & request("Det")&"','"
       sql= sql & RPN &"','Y','"
       sql= sql & request.cookies("fmea.truename") &"','"
       sql= sql & Now() &"')" 
    
       call openDB()
       conn.execute(sql)
       call closeDB()				       
End Select
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript" type="text/javascript">
<!--
function jump() 
{
 if(event.keyCode==13)event.keyCode=9
}
-->
</script>
<title>项目输入</title>
</head>

<body   marginheight="0" topmargin="10" bottommargin="0" leftmargin="0" bgcolor="#f9f9f9">
<form action="Process_Add.asp" method="post">
<table width="66%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center">
<tr bgcolor="#6c9c6c">
<td colspan ="2" height ="30" align ="center">
<font size ="3" color ="ffffff"><% = Request("PjKey") %>评价内容输入</font>
<input name="PjID" type ="hidden" value ="<% = Request("PjID") %>">
<input name="PjKey" type ="hidden" value ="<% = Request("PjKey") %>">
</td></tr>
<tr bgcolor="#FCfCFC">
<td width = "150">工程内容</td>
<td><input name="SubPjName" class="notline" size="40">
</td></tr>
<tr bgcolor="#FCfCFC">
<td width = "150">作业标准书页码</td>
<td><input name="PageNo" class="notline" size="25">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>作业顺序</td>
<td><input name="WorkQueue" class="notline" size="5">
</td></tr>
<tr bgcolor="#FCfCFC">
<td>部品编号</td>
<td><input name="PartsNo" class="notline" size="25"></td></tr>
<tr bgcolor="#FCfCFC">
<td>部品名称</td>
<td><input name="PartsName" class="notline" size="70"></td></tr>
<td bgcolor="#FCfCFC">部品种类</td>
<td bgcolor="#FCfCFC">
<div>
<input name="PartsType" type ="Radio" value ="板金"> 板金
<input name="PartsType" type ="Radio" value ="塑胶"> 塑胶
<input name="PartsType" type ="Radio" value ="轮轴"> 轮轴
<input name="PartsType" type ="Radio" value ="弹簧"> 弹簧
<input name="PartsType" type ="Radio" value ="感光鼓">感光鼓
<input name="PartsType" type ="Radio" value ="组立部件">组立部件
</div><div>
<input name="PartsType" type ="Radio" value ="电线"> 电线
<input name="PartsType" type ="Radio" value ="基板"> 基板
<input name="PartsType" type ="Radio" value ="马达"> 马达
<input name="PartsType" type ="Radio" value ="感应器"> 感应器
<input name="PartsType" type ="Radio" value ="离合器"> 离合器
</div><div>
<input name="PartsType" type ="Radio" value ="轴承"> 轴承
<input name="PartsType" type ="Radio" value ="卡环"> 卡环
<input name="PartsType" type ="Radio" value ="标签"> 标签
<input name="PartsType" type ="Radio" value ="螺丝"> 螺丝
<input name="PartsType" type ="Radio" value ="油"> 油
<input name="PartsType" type ="Radio" value ="耗材"> 耗材
</div>
</td></tr>
<tr bgcolor="#FCfCFC">
<td>作业分类</td>
<td>
<select name="WorkCata" class="notline">
 <option value="打螺丝">打螺丝</option>
 <option value="装入">装入</option>
 <option value="张贴">张贴</option>
 <option value="结线">结线</option>
 <option value="涂布">涂布</option>
 <option value="清扫">清扫</option>
 <option value="设定">设定</option>
 <option value="线处理">线处理</option>
</select>
</td></tr>
<tr bgcolor="#FCfCFC">
<td>作业内容</td>
<td><input name="WorkContent" class="notline" size="100"></td>
</tr>
<tr bgcolor="#FCfCFC">
<td>潜在缺陷模式</td>
<td><input name="NGMode" class="notline" size="100"></td></tr>
<tr bgcolor="#FCfCFC">
<td>潜在缺陷后果</td>
<td><input name="NGEffect" class="notline" size="100"></td></tr>
<tr bgcolor="#FCfCFC">
<td>现行工程设计及过程控制</td>
<td><input name="ProDesign" class="notline" size="100"></td></tr>
<tr bgcolor="#FCfCFC">
<td>严重度Sev</td>
<td>
<input name="Sev" type ="Radio" value ="1"> 1
<input name="Sev" type ="Radio" value ="2"> 2
<input name="Sev" type ="Radio" value ="3"> 3
<input name="Sev" type ="Radio" value ="4"> 4
<input name="Sev" type ="Radio" value ="5"> 5
<input name="Sev" type ="Radio" value ="6"> 6
<input name="Sev" type ="Radio" value ="7"> 7
<input name="Sev" type ="Radio" value ="8"> 8
<input name="Sev" type ="Radio" value ="9"> 9
<input name="Sev" type ="Radio" value ="10"> 10
</td></tr>
<tr bgcolor="#FCfCFC">
<td>频度Occ</td>
<td>
<input name="Occ" type ="Radio" value ="1"> 1
<input name="Occ" type ="Radio" value ="2"> 2
<input name="Occ" type ="Radio" value ="3"> 3
<input name="Occ" type ="Radio" value ="4"> 4
<input name="Occ" type ="Radio" value ="5"> 5
<input name="Occ" type ="Radio" value ="6"> 6
<input name="Occ" type ="Radio" value ="7"> 7
<input name="Occ" type ="Radio" value ="8"> 8
<input name="Occ" type ="Radio" value ="9"> 9
<input name="Occ" type ="Radio" value ="10"> 10
</td></tr>
<tr bgcolor="#FCfCFC">
<td>不易探测度Det</td>
<td>
<input name="Det" type ="Radio" value ="1"> 1
<input name="Det" type ="Radio" value ="2"> 2
<input name="Det" type ="Radio" value ="3"> 3
<input name="Det" type ="Radio" value ="4"> 4
<input name="Det" type ="Radio" value ="5"> 5
<input name="Det" type ="Radio" value ="6"> 6
<input name="Det" type ="Radio" value ="7"> 7
<input name="Det" type ="Radio" value ="8"> 8
<input name="Det" type ="Radio" value ="9"> 9
<input name="Det" type ="Radio" value ="10"> 10
</td>
</tr>
<tr bgcolor="#FCfCFC" height ="30">
<td colspan = "2" align ="center" >
<input name="OK" value="项目输入" type="submit" class="smallInput">
</td></tr>
</table>
</form>
</body>
</html>

