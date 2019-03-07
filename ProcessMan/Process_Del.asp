<!--#include file="../includes/db.asp"-->
<% 
Select Case Request("cat")
      Case 1
          sql = "Insert into t_Fmea_Item_history "
          sql = sql & "(ItmID,PjID,Pjkey,SubPjName,PageNo,WorkQueue,WorkTheme,PartsNo,PartsName,PartsType,WorkCata,"
          sql = sql & "WorkContent,NGMode,NGEffect,ProDesign,Sev,Occ,"
          sql = sql & "Det,RPN,ItmStatus,InUser,InTime)"
          sql = sql & " Select ItmID, PjID ,Pjkey,SubPjName,PageNo,WorkQueue,WorkTheme,PartsNo,"
          sql = sql & "PartsName,PartsType,WorkCata,WorkContent,NGMode,NGEffect,"
          sql = sql & "ProDesign,Sev,Occ,Det,RPN,'N',InUser,InTime"
          sql = sql & " from t_Fmea_Item Where ItmID = '"& request("ItmID") & "'"
            
          SqlStr_c ="Update t_Fmea_Item set "
          SqlStr_c = SqlStr_c & "ItmStatus = 'N',"
          SqlStr_c = SqlStr_c & "InUser = '"& request.cookies("fmea.truename") & "',"
          SqlStr_c = SqlStr_c & "InTime = '"& now() & "'"
          SqlStr_c = SqlStr_c & " Where ItmID = '"& request("ItmID") & "'"
          
          call openDB()
          conn.BeginTrans
          conn.execute(sql)
          conn.execute(SqlStr_c)           
          conn.CommitTrans
          call closeDB()
    Case 2
         sql_ad = "Insert into t_Fmea_Advice_history "
         sql_ad = sql_ad & "(AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
         sql_ad = sql_ad & "ActionContent,ResultS,ResultO,ResultD,"
         sql_ad = sql_ad & "ResultRPN,AdvStatus,InUser,InTime)"
         sql_ad = sql_ad & " Select AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
         sql_ad = sql_ad & "ActionContent,ResultS,ResultO,ResultD,"
         sql_ad = sql_ad & "ResultRPN,'N',InUser,InTime"
         sql_ad = sql_ad & " from t_Fmea_Advice Where ItmID = '"& request("ItmID") & "'"        
 
         sql = "Insert into t_Fmea_Item_history "
         sql = sql & "(ItmID,PjID,Pjkey,SubPjName,PageNo,WorkQueue,WorkTheme,PartsNo,PartsName,PartsType,WorkCata,"
         sql = sql & "WorkContent,NGMode,NGEffect,ProDesign,Sev,Occ,"
         sql = sql & "Det,RPN,ItmStatus,InUser,InTime)"
         sql = sql & " Select ItmID,PjID,Pjkey,SubPjName,PageNo,WorkQueue,WorkTheme,PartsNo,"
         sql = sql & "PartsName,PartsType,WorkCata,WorkContent,NGMode,NGEffect,"
         sql = sql & "ProDesign,Sev,Occ,Det,RPN,'DL',InUser,InTime"
         sql = sql & " from t_Fmea_Item Where ItmID = '"& request("ItmID") & "'"
                 
         SqlStr_c = "Delete t_Fmea_Item Where ItmID = '"& request("ItmID") & "'"

         call openDB()
         conn.BeginTrans
         conn.execute(sql_ad)
         conn.execute(sql)
         conn.execute(SqlStr_c)           
         conn.CommitTrans
call closeDB()

End Select
          
Response.Redirect "Process_Browse.asp?page="&request("page")

%>