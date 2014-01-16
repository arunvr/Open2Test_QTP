'###############################  KEYWORD DRIVEN FRAMEWORK   #####################################
'Project Name		: Windows Framework
'Author		       	: Open2Test
'Version	    	: V 2.1
'Date of Creation	: 12-Mar-2013
'######################################  Driver Function  ##################################################

'#################################################################################################
'Function name 	: 		Keyword_Win
'Description             : This is the main function which interprets the keywords and performs the desired actions. All the keywords 
'                                      used in the datatable are processed in this function.
'Parameters           : Keyword in the 2nd column of the data table
'Assumptions     	: The Automation Script is present in the Local Sheet of QTP.
'#################################################################################################
'The following function is for 'Context' keyword.
'#################################################################################################
Function Keyword_win(initial)
   	 If htmlreport = "1" Then
		Call Update_Log(MAIN_FOLDER, g_sFileName,"executed")' calling function update log to create an execution log in HTML file
	 End If	
   On error resume next
	If initial = "context" Then
	   Call Func_Context_Win(arrAction,intRowCount)
	   Exit Function
