<html>
<!--
'''�������ܣ���ѯ���û���Ϣ 
'''����˵����strAdmin��������˻���Password�����ʻ����룻Domain����������� 
''' 
''' �ο����ϣ�http://www.experts-exchange.com/Web_Development/Web_Languages-Standards/ASP/Q_23947512.html 
''' ������Query LDAP For Existing User   Classic ASP �� www.experts-exchange.com 
function getADUserData(strAdmin,strPassword,Domain,userName) 

	If  AuthenticateUser(strAdmin,strPassword,Domain)=False Then '�������Ա��֤ʧ�����˳��˹��� 
           Response.Write "��֤ʧ�ܡ�" 
           Exit function 
	End If 

	Dim Conn, strRS, RS, strConn  
	Set Conn = Server.CreateObject("ADODB.Connection")  
	Set RS = Server.CreateObject("ADODB.Recordset")  
	  
	Conn.Provider = "ADsDSOObject"  
	Conn.Properties("User ID") = strAdmin 
	Conn.Properties("Password") = strPassword 
	Conn.Properties("Encrypt Password") = True 
	strConn = "Active Directory Provider"  
	Conn.Open strConn , strAdmin, strPassword 

	strRS = "SELECT name FROM 'LDAP://" & Domain & "' WHERE sAMAccountName = '"&userName&"' ORDER by name "  
	RS.Open strRS, Conn,1,1     

	While RS.EOF = False  
	getADUserData=rs.Fields("name") 
	RS.MoveNext  
	Wend 

End function  


'''�������ܣ���֤���û��˺����� 
'''����˵����UserName�����˻���Password�����ʻ����룻Domain����������� 
'''���أ� �û��������˺�������ȷ�򷵻�True�����򷵻�False�� 
''' �ο����ϣ�http://stackoverflow.com/questions/3894835/ldap-asp-classic-adodb-2147217865-using-ldap-to-talk-to-active-directory 
''' ������LDAP + ASP Classic + ADODB  �� stackoverflow 
''' ������Getting Started with ASP for ADSI �� ΢��MSDN 

function AuthenticateUser(UserName, Password, Domain) 
	dim strUser  
	' assume failure  
	AuthenticateUser = false   
	strUser = UserName 
	strPassword = Password 
	strQuery = "SELECT cn FROM 'LDAP://" & Domain & "' WHERE objectClass='*' "  
	set oConn = server.CreateObject("ADODB.Connection")  
	oConn.Provider = "ADsDSOOBJECT"  
	oConn.Properties("User ID") = strUser  
	oConn.Properties("Password") = strPassword  
	oConn.Properties("Encrypt Password") = true  
	oConn.open "DS Query", strUser, strPassword 
	set cmd = server.CreateObject("ADODB.Command")  
	set cmd.ActiveConnection = oConn  
	cmd.CommandText = strQuery  
	on error resume next  
	set oRS = cmd.Execute  
	if oRS.bof or oRS.eof then 
	AuthenticateUser =  false  
	else     
	AuthenticateUser = True  
	end if  
	set oRS = nothing  
	set oConn = nothing   
end function 

-->



<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<script   language=javascript   RUNAT="SERVER">   
function   logonDoADLogon(p_strDomain,   p_strUserID,   p_strPWD)   
{               //return   true;   
var   f_oIADS,   f_oUser,   f_oContainer;   
var   f_blnRet   =   true;   
    
try   
{   
var   f_oIADS   =   GetObject('WinNT:');   
f_oContainer   =   f_oIADS.OpenDSObject('WinNT://'   +   p_strDomain,   p_strDomain   +   "\\"   +   p_strUserID,   p_strPWD,   0);   
    
delete   f_oContainer;   
delete   f_oIADS;   
}   
catch   (e)   
{   
return   false;   
}     
    
try   
{   
var   objUser   =   GetObject("WinNT://"   +   p_strDomain   +   "/"   +   p_strUserID   );   
delete   objUser;   
}   
catch(e){   
return   false;   
}       
    
return   true;   
}   
</script>   
    

<title>����FMEA����ϵͳ</title>
</head>

<body>


<%   
'id=trim(request("id"))   
'pwd=trim(request("pwd"))
id = "cn015429"
pwd = "Zjf12345678" 
domain ="kmcn.local"  
'domain=trim(request("domain"))   
if     logonDoADLogon(domain,id,pwd)   then   
response.write   "True"   
else   
response.write   "False"   
end   if   
%>


<%


'''���ã� 
'ADUserName=getADUserData("��֪�˻�","��֪����","kmcn.local","Ҫ��ѯ�����˻�")

dim f
for each f in request.ServerVariables
response.write(request.ServerVariables(f)&"<hr>")
next


%>
</body>
</html>