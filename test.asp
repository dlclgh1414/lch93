﻿<!--#include virtual="common/DBHelper.asp"-->
<!--#include virtual="common/json_2.0.4.asp"-->
<!--#include virtual="common/JSON_UTIL_0.1.1.asp"-->

<%
	Response.Charset = "utf-8"
	
	set  DBHelper = new clsDBHelper
	
			strQuery = 					  " SELECT L_LAW_IDX, L_LAWS, L_ITEM, L_LAW, L_REMARK , UL_STS, UL_STS_UPGRADE_PLAN "
			strQuery = strQuery & 	" FROM TB_LAW LEFT JOIN TB_USER_LAW ON L_LAW_IDX = UL_LAW_IDX AND UL_USER_IDX = 1"
			'strQuery = strQuery & 	" WHERE L_CHK_GUBUN = 1 AND L_LAW IS NOT NULL"
'			strQuery = strQuery & 	" ORDER BY CONVERT(INT,SUBSTRING(L_LAWS,2,PATINDEX('%��',L_LAWS)-2)), LEFT(L_ITEM,PATINDEX('%-%',L_ITEM)+1),CONVERT(int,REPLACE(L_ITEM,'-','')) "
			
			Set ListRs = DBHelper.ExecSQLReturnRS(strQuery,Nothing, Nothing)
			
			
			
			If ListRs.Recordcount > 0 Then 
				ListRs.movefirst
				ListRec = ListRs.GetString(,,"¿","†")    
				List_rows = split(ListRec,"†")    
			End if    
			ListRs.Close
			Set ListRs = Nothing		


			response.write ListRec
%>

