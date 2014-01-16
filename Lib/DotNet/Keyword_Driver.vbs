'###############################  KEYWORD DRIVEN FRAMEWORK   #####################################
'Project Name		: Web Framework
'Author		       	: Open2Test
'Version	    	: V 2.0
'Date of Creation	: 31-May-2012
'######################################  Driver Function  ################################################
Option Explicit     'To enforce variable declaration
Dim arrTech 		'This variable stores the technology  , used in Func_TechInitialize()
Dim vTech  			'This variable is used in Func_TechnInitialize()
Dim iTech 			'This variable is used for looping
Dim arrTechList 	'This variable is used to stored the technology 
Dim FrameName  	    'This variable is used to store the page name 
Dim parent          'This variable is used to store the current context
Dim curParent 	    'This variable is used to store the current parent
Dim parChild	    'This variable is used to store last item in 4th Column of Datasheet
Dim parParent 		'This variable is used to store the type of the parent element, used in SAP framework
Dim propertyVal     'This variable is used to store the current property to be used
Dim arrAction	  	'This variable is used to store the object type and name
Dim arrKeyValue		'This variable is used to store the values present in 4th Column of DataSheet
Dim arrKeyIndex		'This variable is used to store the split array of 4th Column of DataSheet
Dim childCount	 	'This variable is used to store the child item count
Dim index		   	'This variable is used to store the object index
Dim initial			'This variable is used to store the value in 2nd Column of DataSheet
Dim propFound		'This variable is used to store the flag value in Checking
Dim VarName		    'This variable is used to store the returned values from the application 
Dim propName	    'This variable is used to store the current property which is to be used
Dim propSplit	    'This variable is used to store the array after Split operation
Dim strParam        'This variable is used to store the operation to be performed in table operations
Dim strParam1	 	'This variable is used to store the operation to be performed in table operations
Dim object          'This variable contains the current object
Dim keyword		  	'This variable is used to indicate if a condition has passed or not 
Dim intCounter	  	'This variable is used for storing loop count in 'For' loops
Dim newContext      'This variable is used to indicate if new context is set
Dim intRowCount     'This variable is used as a counter to loop through all the data table rows
Dim intDataCounter  'This variable is used to store the iteration count for looping
Dim strCellData     'This variable is used to pass values to 'GetValue' function
Dim strIexplorePath	'This variable is used to store the Path of the internet explorer
Dim inti			'Used for looping
Dim intj   	  		'Used for looping
Dim intSheet		'Used to check whether Keyword Script is present in the Local Sheet
Dim objName			'This variable is used to store the Object Name
Dim objPerform		'This variable is used to store the value present in the fourth Column
Dim aPosition 		'This variable stores the position where match value is found in arrObj(1)
Dim aMatch 			'This variable stores the match value found in arrObj(1) 
Dim dbConn 			'This variable is used to store the database connection object
Dim dbRs 			'This variable holds the result of the database operation performed
Dim strSQL	   		'This variable stores the query to be executed
Dim strReplace 		'This variable stores the Environment value of aMatch
Dim connectionString'This variable stores the connection string for Database
Dim dbUID			'This variable is used to store the user name to connect to database
Dim dbPWD			'This variable is used to store the password to connect to database 
Dim dbServer		'This variable is used to store the database server name
Dim dbHost			'This variable is used to store the database host name
Dim dbDRIVER		'This variable stores the database driver name
Dim curObjClassName	'This variable stores the current object name
Dim arrObjchk		'This variable is used to store the object type, name and the type of match
Dim irowNum			'This variable stores the data table row count
Dim Rep_value		'This variable is used to replace the variable name with the actual value while using the 'Report' keyword
Dim LocalSheet     'This variable stores values in the Action 1 sheet of datatable
Dim strimportdatapath 'This variable stores the value of 3rd column of  datatable '************************
Dim initialloop '*******************
Dim errStr1 'This variable used to stores the error descriptions
Dim errStr2	'This variable used to stores the error descriptions
Dim arrObj  'This variable stores the array of object names
Dim arrTableindex ''This vairable stores the 4th column of datatable, which is in use for the func_tablesearch
'Loop variables
Dim lpflag 'This variable is used as initialization of a variable 
Dim LoopNo 'This variable is used for incremental value
Dim EloopNo  'This variable is used for incremental value
Dim Loopcnt() 'This variable stores the iteration of the loop
Dim Loopind() 'This variable is used for loop functions
Dim LpStrtRow() 'This variable stores the datatable row count where the loop should start with
'Condition Variables
Dim cflag  'This variable is used for condition function 
Dim ecflag 'This variable is used for end condition function
'log variables
Dim MAIN_FOLDER  'This variable stores the path where the execution log to be stored
Dim g_sFileName 'This variable stores the name of the test with execution time stamp
Dim htmlreport 'This initiialization variable is used for log function
Dim arrsearchcriteria 'Stores the set of search criteria in an array          
Dim arreachsearch 'Stores each search criteria in an array                
Dim arrOutColsvars'Stores the variable names in an array to return the column number 
Dim searchtablereturn 'Stores the retun value from the function            
Dim arrsearchtablereturn 'Stores the Return value from the function  
Environment.Value("ErrorLog")=0  '***************************************
Environment.Value("intEndRow")=0   'initialized and takes the end row number from datatable , where which the debug should end
Environment.Value("iperform")=0  'Environment variable  initialized to store the screen shot of perform function under customization
Environment.Value("icontext")=0   'Environment variable  initialized to store the  screen shot of context function under customization
Environment.Value("icheck")=0   'Environment variable  initialized to store the screen shot of check function under customization
Environment("LogFile")=0				'Environment variable initialized used for log funtions
Environment("PrintOption")=1 			'Environment variable initialized for the printlog file in log functions
'#################################################################################################
'Function name 		: Keyword_Driver
'Description        : This function is used to call the main framework
'Parameters       	: NA
'Assumptions     	: The Automation Script is present in the Local Sheet of QTP.
'#################################################################################################
'The following function is for  Keyword_Driver 
'#################################################################################################
Public Function Keyword_Driver()
   On error resume next
	Setting("DefaultTimeout")= 2000
	Call Func_TechInitialize()   
    intRowCount = 1
	 If intDataCounter = empty Then
		 intDataCounter = 1
	End If
    intSheet = 0
	 While (intRowCount <= DataTable.LocalSheet.GetRowCount)
		DataTable.LocalSheet.SetCurrentRow(intRowCount)
		Dim x
		x = DataTable.Value(1,dtLocalSheet)
		    If (trim(LCase(DataTable.Value(1,dtLocalSheet))) = "r") Then
				 Call Keyword_Call()
			End If
			Call Func_Error
        intRowCount = intRowCount + 1
		intSheet = 1
	 Wend
	 If intSheet = 0 Then
		 Reporter.ReportEvent micFail, "Keyword Script should be present in the Local Sheet", "Script is not present in the Local Sheet, Please verify the  Data table"
	 End If
	If cflag <> ecflag Then
		Reporter.ReportEvent micFail, "Condition Construct Exception Occured.", "Please check that every Condition keyword has its corresponding endCondition keyword" 
	End If
	If LoopNo <> EloopNo Then
		Reporter.ReportEvent micFail, "Loop Construct Exception Occured.", "Please check that every Loop keyword has its corresponding endLoop keyword" 
	End If
    If htmlreport = "1" Then
		 Call Update_log(MAIN_FOLDER, g_sFileName,"finish")
	 End If
	 Call Func_Error()
 End Function
'#################################################################################################

'#################################################################################################
' Function Name 	: Keyword_Call
'Description        : This is the main function which interprets the keywords and performs the 
'					  				desired actions. All the keywords used in the datatable are processed in this function
'Parameters       	: NA
'Assumptions     	: The Automation Script is present in the Local Sheet of QTP.
'#################################################################################################
'The following function is called internally
'#################################################################################################
'#################################################################################################
Public Function keyword_Call()
	If htmlreport = "1" Then
		Call Update_Log(MAIN_FOLDER, g_sFileName,"executed")
	 End If	
    initial = LCase(Trim(DataTable.Value(2,dtLocalSheet))) 'Storing the generic action type to be performed.
   If (DataTable.Value(3,dtLocalSheet) <> "") Then
		objName = Cstr(Trim(Lcase(DataTable.Value(3,dtLocalSheet))))
		arrAction = (Split(objName, ";", -1, 1)) 'splits the value into object type and name
        If initial <> "arith" Then
			For  inti = 0 to Ubound(arrAction)
				arrAction(inti) = GetValue((Trim(arrAction(inti))))'Getting the values stored in variables.
			Next
		End If
	End If
     If (DataTable.Value(4,dtLocalSheet) <> "") Then 'Checking if any value is presenting the 4th Column 
		objPerform = CStr(Trim(DataTable.Value(4,dtLocalSheet)))
		arrKeyIndex = split(objPerform, ":", -1,1)'
		arrKeyValue = Split(objPerform, ":", -1, 1)
		arrTableindex=Split(objPerform, ":", 2, 1)'Splitting the value into specific action type and value to be used with row and column values for table operations.
		For inti = 0 to Ubound(arrKeyValue)
			arrKeyValue(inti) = GetValue((Trim(arrKeyValue(inti))))'Getting the values stored in variables.
		Next
		For  inti = 0 to Ubound(arrKeyIndex)
			arrKeyIndex(inti) = GetValue((Trim(arrKeyIndex(inti))))'Getting the values stored in variables.
		Next
	End If
	Select Case LCase(initial) 'To perform user defined keywords
		Case "importdata"
            Call func_ImportData(objName)
		Case "screencaptureoption"			
			Call Func_ScreencaptureOption()
		Case "log"
			htmlreport = "1"
		   Call createlog(objName, objPerform)
        Case "debug"
			Call Func_Debug()
		Case "capturedata"
			Call Func_CaptureData()
		Case "exit"
			On Error GoTo 0
			irowNum = Datatable.LocalSheet.GetRowCount
			intRowCount = irowNum
			Datatable.LocalSheet.SetCurrentRow(intRowCount) 
			Reporter.ReportEvent micDone,"Exit Test Script action called","Test was exited on user request"			
			ecflag = cflag
			EloopNo = LoopNo
		Case "convert"
            Environment.Value(lcase(arrkeyIndex(1)))= Func_convert()
        Case "assignvalue"
			Environment.Value(lcase(arraction(0))) = arrAction(1)
		Case "report"
            Rep_Value = Rep_Variable(objName, ":")
            Call Func_Report()
			Rep_Value = ""
		Case "msgbox"
			msgbox arrAction(0)
		Case "launchapp"			
            Select Case UCase(Trim(DataTable.Value(4,dtLocalSheet))) 'To launch the application in different browser
				Case "IE"
						strIexplorePath = "C:\Program Files\Internet Explorer\iexplore.exe"
				Case "SAFARI"
						strIexplorePath = "C:\Program Files\Safari\Safari.exe"
				Case "FIREFOX"			
						strIexplorePath  = "C:\Program Files\Mozilla Firefox\firefox.exe"
				Case else
					strIexplorePath = "C:\Program Files\Internet Explorer\iexplore.exe"
				End Select
            If (DataTable.Value(3,dtLocalSheet) <> "") Then
				strCellData = Getvalue(datatable.Value(3,dtLocalSheet))
				If ((Right(strCellData,4)=".exe") or ((Right(strCellData,4)=".lnk"))) Then
					SystemUtil.Run strCellData,"","",""
				Else
					SystemUtil.Run strIexplorePath, strCellData, "", "open"
				End if	
			Else
				SystemUtil.Run strIexplorePath, Environment.Value("LaunchApp"), "", "open"
			End If 'End of Launch app
		Case "presskey"
			Call Func_presskey()
		Case "callnestedaction" '
			Call Func_CommonFunctions("callnestedaction")
		Case "sendemail"
			Call Func_emailcore()
        Case "arith"
			Rep_Value = Rep_Variable(objName, "arith")
			Environment.Value(lcase(arrkeyIndex(0))) = eval(Rep_Value)
			Rep_Value = ""
        Case "callaction" 'common functions
			Call Func_CallAction()
		Case "callfunction"
			If isArray(arrAction) = False then
				Reporter.ReportEvent micFail,"User defined function exception occurred.", "Please provide the User Defined Function name in the keyword test step " &intRowcount
            elseif Ubound(arrAction) = 1 then
				Environment.Value(lcase(arrAction(1))) = Func_FunctionCall()
			Else
				Call Func_FunctionCall()
			End if
		Case "function"
			 If arrAction(0) = "folder" Then
				Call Func_Folder(objPerform)
			elseif arrAction(0) = "file" then
				Call Func_File(objPerform)
            End If
       Case "loop"
			If (instr(1,objName,";",1)) then
                Call Func_loop()	
			else
				LoopNo = LoopNo + 1
                If lpflag = 0 Then
					lpflag=1
					ReDim preserve Loopcnt (0)
					ReDim preserve Loopind (0)
					ReDim preserve LpStrtRow(0)
					intdatacounter = 1	
					If objName <> "" Then
						Loopcnt(Ubound(Loopcnt))=arrAction(0)
					else  
						Loopcnt(Ubound(Loopcnt)) = Datatable.GlobalSheet.getrowcount
						If Loopcnt(Ubound(Loopcnt)) = 0 Then
								Loopcnt(Ubound(Loopcnt)) = 1
						End If
                    End If					
					LpStrtRow(Ubound(LpStrtRow)) = intRowCount
					Loopind(Ubound(Loopind))=0
				Else
					lpflag = 1
					ReDim preserve Loopcnt (Ubound(Loopcnt)+1)
					ReDim preserve Loopind (Ubound(Loopind)+1)
					ReDim preserve LpStrtRow(Ubound(LpStrtRow) + 1)
					Loopcnt(Ubound(Loopcnt))=arrAction(0)
					LpStrtRow(Ubound(LpStrtRow)) = intRowCount
					Loopind(Ubound(Loopind))=0
				End if
			End If  
		Case "endloop"
				Dim lpres
                If Loopind(Ubound(Loopind))=0 Then
					EloopNo = EloopNo+ 1
                End if
				lpres = Func_Eloop
				If lpres <> 0 Then
					intRowCount = lpres
				End If
		Case "wait" 'common functions
			Wait CInt(arrAction(0))
		Case "strsearch"
			Environment.Value(lcase(arrkeyIndex(0))) = Func_StringOperations("strsearch")
		Case "strconcat"
			Rep_Value = Rep_Variable(objName, ";")
			Environment.Value(lcase(arrkeyIndex(0))) = Func_StringOperations("strconcat")
		Case "strreplace"
			Rep_Value = Rep_Variable(objName, ";")
			Environment.Value(lcase(arrkeyIndex(0))) = Func_StringOperations("strreplace")
		Case "strreverse"
			Rep_Value = Rep_Variable(objName,";")
			Environment.Value(LCase(arrkeyIndex(0))) = Func_StringOperations("strreverse")
		Case "condition"
			If objPerform <> "" Then
				Call Func_Condition(intRowCount)
			Else
				cflag = cflag + 1
				Dim res1 'temporary vairable
				res1 =  Func_Condition1
				If res1<>0 Then
					introwcount = res1
				End If
			End if
		Case "endcondition"
			ecflag = ecflag + 1
		Case "label"		
		Case "jumpto"
			Dim resj
			resj = Func_goto(arrAction(0))
			If (resj <> 0) Then
				intRowCount = resj
				Datatable.LocalSheet.SetCurrentRow(intRowCount-1)
			Else
				Datatable.LocalSheet.SetCurrentRow(intRowCount)
			End if
        Case "screencapture"	
			Call Func_CaptureScreenshot("test",intRowCount)
		Case "exportxml"
			Call Func_ExportXML(objName,Datatable.Value(4,dtLocalSheet))
		Case "deletexml"
			Call Func_DeleteXML(objName)
		Case "getxmlvalue"
			Call Func_GetXMLValue(objName)
		Case "setxmlvalue"
			Call Func_SetXMLValue(objName)
        Case "getparameter"
    		  Environment.Value(Lcase(arrAction(0)))=Parameter(arrAction(1)) 
		Case Else	'If initial is other than launchapp, context, presskey, arith, call action, brokenlinks
			Execute "Call Keyword_" & arrTech(vTech)&"(initial)"
			If err.number = -2147220990 Or err.number = 424 Or err.number = vbObjectError Then
				For iTech=0 to Ubound(arrTech)
					If vTech <> iTech Then
						Err.clear
						Execute "Call Keyword_" & arrTech(iTech)&"(initial)"
					End If					
                    If err.number <> -2147220990 and err.number <> 424 and err.number <> vbObjectError Then
						errStr1 = ""
						errStr2 = ""
						err.clear
						vTech = iTech
						Exit for
					End If
				Next
				If errStr1<>"" then
					Reporter.ReportEvent micFail, errStr1,errStr2
					errStr1 = ""
					errStr2 = ""
				End if
			End if
	End Select
    If isArray(arrAction) Then
		Erase arrAction
	End If
    If isArray(arrkeyIndex) Then
        Erase arrkeyIndex
	End if
	If isArray(arrkeyValue) Then
		Erase arrKeyValue
	End if
	objPerform = ""
	objName = ""
 End Function ' End of Main Function
'#################################################################################################################

'#################################################################################################################
'Function name 		: GetValue
'Description        : This function is used to retrieve the value from any variable.
'Parameters       	: 'strCellData'containing the value passed from other functions goes as a parameter.
'Assumptions     	: None
'#################################################################################################################
'The following function is for retrieving a value from a variable
'#################################################################################################################
Public Function GetValue(ByRef strCellData)
	Dim arrSplitCheckData 'Stores the elements after the  value is split with "_" delimiter
	Dim strParamName      'Stores the 2nd element of array 'arrSplitCheckData' 
	If Instr(1,strCellData,"#",1) = 1 Then
		strCellData = Environment.Value(lcase(Right(strCellData,Len(strCellData)-1))) 
    ElseIf Lcase(strCellData) = "blank" Then
		strCellData = ""	
	Else
		arrSplitCheckData = Split(strCellData,"_",2,1)
		If UBound(arrSplitCheckData) > 0 Then
			Select Case LCase(arrSplitCheckData(0))  'Retrieving the values of any variable defined through parameter, environment, datatable,etc
				Case "p"
					strCellData = Parameter(arrSplitCheckData(1))
				Case "env"
					strCellData = Environment.Value(lcase(arrSplitCheckData(1)))
				Case "dt"
					 If Datatable.value(5,dtLocalSheet) <> "" Then
							Dim strPath     'stores the path with .xls 
							Dim strDataPath 'locates and stores the full path 
							Dim strSheetName'stores the sheet name to be imported
							Dim strTestCase  ' Stores the value in 5th column of datatable
							Dim strSheet	'stores the sheet name of the excel sheet
							strTestCase = CStr(Trim(Datatable.value(5,dtLocalSheet)))
							strSheetName = Split(strTestCase,";",2,1)
							If InStr(1,strSheetName(0),".xls") = 0 Then	   'check if .xls is given in the path provided.
							   strPath     = strSheetName(0) & ".xls"		   'adding .xls to the Path
							Else
							   strPath     = strSheetName(0)
							End If
							If Ubound(strSheetName) = 1 Then
                                If (Asc(strSheetName(1)) <58) Then
											If (Asc(strSheetName(1))>47) Then
												strSheet = Cint(Trim(strSheetName(1)))
											End If										
								 Else
										strSheet = Cstr(strSheetName(1))
								 End If
							Else
								strSheet = 1	
							End If
							If InStr(1,strSheetName(0),":\") <> 0 Then
								strDataPath = strPath	           'storing the full path in strDataPath
							Else
								strDataPath = Pathfinder.Locate(strPath) 'locating and storing the full path in strDataPath
							End If
							If strDataPath <> "" Then	'Check if the sheet is present in the given path
								DataTable.ImportSheet strDataPath,strSheet, 1  'If present import the data into Action1 sheet
							Else
								Reporter.ReportEvent micFail,"Incorrect Input","The file" & " " & strPath & " " & "is not present in the attachment for the current test. Please check the value of the row " & intRowCount
							End If						 
					 End If
					DataTable.SetCurrentRow(intDataCounter)
                    			strParamName = arrSplitCheckData(1)
					strCellData =Datatable.Value(strParamName,"Global")
					DataTable.SetCurrentRow(intRowCount)
			End Select
		End If
	End If
	GetValue = strCellData
End Function
'#################################################################################################

'#################################################################################################
'Function name 		: Rep_Variable
'Description        : This function is used to replace the varible 
'Parameters       	: 1.Value to be replaced  2.delimiter ( start  position  of replacing the character) are passed as parameter  
'Assumptions     	: None
'#################################################################################################

'#################################################################################################
Public Function Rep_Variable(Rep_actval, dlim)
	Dim var_pos 'stores the position of delimiter '#" in the given string
	Dim dlim_pos 'stores the position of delimiter from the variable 'dlim' in the given string
	Dim var_name 'variable with the environment value from datatable
	var_pos = 1
	If dlim = "arith" Then
		If instr(1,Rep_actval,"+") Then
			dlim = "+"
		elseif instr(1,Rep_actval,"-") then
			dlim = "-"
		elseif instr(1,Rep_actval,"*") then
			dlim = "*"
		elseif instr(1,Rep_actval,"/") then
			dlim = "/"
		elseif instr(1,Rep_actval,"^") then
			dlim = "^"
		End If
	End If		
	While instr(var_pos,Rep_actval,"#")
		var_pos = instr(var_pos,Rep_actval,"#")
		dlim_pos = instr(var_pos,Rep_actval,dlim)
		If dlim_pos = 0 Then
			dlim_pos = len(Rep_actval)+1
		End If
		var_name = lcase(mid(Rep_actval,var_pos+1,(dlim_pos - var_pos-1)))		
		Rep_actval = Replace(Rep_actval,var_name,Environment.Value(lcase(var_name)),1,1,1)
        var_pos = var_pos+1
	Wend
	Rep_Variable = Replace(Rep_actval,"#","")
End Function

'#####################################################################################################

'#####################################################################################################
'Function name 		: Func_Error()
'Description        : This function is used to to give the description of the error
'Parameters       	: None
'Assumptions     	: None
'####################################################################################################
'The following function is for retrieving a value from a variable
'####################################################################################################
Function Func_Error()
   Dim strError	'Stores the value present in fifth Column of current row in local sheet 
	If Err.Number <> 0 Then
        If Err.Description <> "" Then
		Reporter.ReportEvent micFail,"ERROR -Occurred at Line :  "&Datatable.LocalSheet.GetCurrentRow,  "Error Description : "& Err.Description
		End If
        Err.Clear
	End If
On error resume next
	If DataTable.Value(5,dtLocalSheet)<> "" Then
		If err.number <> -2147220909 Then
        strError = Datatable.Value(5,dtLocalSheet) 'Storing the value preset in the fifth Column	
		If Lcase(strError) = "onfailureexit" And keyword = 1 Then ' Checking if the checkpoint failed or not
            ExitTest
		End If
		End If
	Err.Clear
	End If
End Function
'#####################################################################################################

'#####################################################################################################
'Function name 		: Func_CheckValueReturn
'Description        : This function is used to store the output of the check function
'Parameters       	: row count of the datatable and variable keyword value is given as parameter. Variable keyword is assigned as '1' or '0''
'									according to status of the result from check function
'Assumptions     	: None
'####################################################################################################
'
'####################################################################################################
Function Func_CheckValueReturn(intRowCount,keyword)
   Dim checkval  'array used to store the value of  splitting checkvalspit
   Dim checkvalsplit 'stores the 5th column value of the datatable
	checkvalsplit = Datatable.Value(5,dtLocalSheet)
	checkval = Split(checkvalsplit, ":", -1, 1)
   If keyword = 1  Then
		Environment(checkval(0)) = "Pass" 'stores the pass result , when the acion, check is passed
   	elseif keyword = 0 then
		Environment(checkval(0)) = "Fail"   ' stores the fail result, when the action fails
		 If checkval(1) = "onerror" Then
			If checkval(2) = "exit" Then
				On Error GoTo 0
				irowNum = Datatable.LocalSheet.GetRowCount
				intRowCount = irowNum
				Datatable.LocalSheet.SetCurrentRow(intRowCount) 
				ExitTest
				Exit Function
			End If
    End If
	End If
End Function
'#####################################################################################################

'#####################################################################################################
'Function name 	  : `ObjectSet
'Description      : This function sets the objects when descriptive programming is used.
'Parameters       : arrObjchk is an array of object names and intRowCount is the current Row number of the datatable
'Assumptions      : NA
'#####################################################################################################
'The following function is called Internally 
'#####################################################################################################
Function Func_DescriptiveObjectSet( arrObjchk,intRowCount)
   Dim arrDPCheck 'array to store objects
   Dim arrDP			'array to store values when more than one property is available
   Dim arrDPVal 	'variable stores ubound value of arrDP
   Dim arrDPLoop 'variable used for looping
   Dim arrDPRE		'array variable, stores the property and value of the object
   Dim ODesc			'Object variable
   Dim arrDPValCheck 	'array variable stores the object name and properties when descriptive property of an object in mentioned in datatable	
   		For inti = 0 to (Func_RegExpMatch ("##\w*##",arrObjchk(1),aPosition,aMatch) - 1)
			strReplace = Environment.Value(Replace(aMatch(inti),"##","",1,-1,1))
			arrObjchk(1) = Func_gfRegExpReplace(aMatch(inti), arrObjchk(1), strReplace)
		Next    	
	arrDPCheck = (Split(arrObjchk(1),",",-1,1))
	arrDPValCheck = (Split(arrObjchk(1),":=",-1,1))
	If  Ubound(arrDPValCheck) <> 0 Then
	If Ubound(arrDPCheck) <> 0 or  Ubound(arrDPCheck) = 0 Then
		arrDP =  (Split(arrObjchk(1),",",-1,1))
		arrDPVal = Ubound(arrDP)
		Set ODesc = Description.Create()
		For arrDPLoop = 0 to arrDPVal
			arrDPRE = (Split(arrDP(arrDPLoop),":=",-1,1))
				Call GetValue(arrDPRE(1))
				ODesc(arrDPRE(0)).Value = arrDPRE(1)
		Next
        		Set arrObjchk(1) = ODesc
	End If
	End If
End Function
'#####################################################################################################

'#################################################################################################
'Function name 		: Func_Folder
'Description        : If the user is working with Folders by using FSO 
'Parameters       	: 1.The details to be used while using FSO.                                
'Assumptions     	: NA
'#################################################################################################
'The following function is for 'Function' Keyword
'#################################################################################################
Function Func_Folder(pCellData)
	Dim arrFolderpath 	'Stores the elements of the folder path separated by delimiter "\"
	Dim intFolderloc	'Stores the element number of the Folder Name
	Dim DestFolder		'Stores the Destination Folder 
	Dim Foldername		'Stores the Folder Name
	Dim oFSO					'Stores the Created object
	Dim arrCellData				'Stores the details of the operation to be performed
	Dim oFolder						'Stores the details of Object created
	arrCellData =split(pCellData,";",-1,1)
	arrFolderpath=(split(arrCellData(1),"\",-1,1))
	intFolderloc=UBound(arrFolderpath)
	Foldername=arrFolderpath(intFolderloc)'Storing the Folder Name
    Set oFSO = CreateObject("Scripting.FileSystemObject")'Creating a FSO object
	Select Case LCase(arrCellData(0))'Selecting the specific action to be performed
		Case "create"
			If Not oFSO.FolderExists(arrCellData(1)) Then 'Checking if folder already exists
				Set oFolder = oFSO.CreateFolder(arrCellData(1))'Creating the Folder
				Reporter.ReportEvent micDone,"Folder should be created" ,"Folder Created at Path: " & (arrCellData(1))
			Else		
				Set oFolder = oFSO.GetFolder(arrCellData(1))
				Reporter.ReportEvent micDone, "Folder should be created","Folder Already Exists at Path: " & (arrCellData(1))
			End If
		Case "delete"	
			If Not oFSO.FolderExists(arrCellData(1)) Then 'Checking if folder already exists
				Reporter.ReportEvent micFail, "Keyword Check at Line no - " & intRowCount,"Folder does not Exist at Path: " & (arrCellData(1))
			Else
				oFSO.DeleteFolder(arrCellData(1))'Deleting te Folder
				Reporter.ReportEvent micDone, "Folder should be deleted","Folder is deleted at Path:" & (arrCellData(1)) 
			End If
		Case "copy"
			If Not oFSO.FolderExists(arrCellData(1)) Then 'Checking if folder already exists
				Reporter.ReportEvent micFail, "Keyword Check at Line no - " & intRowCount  &vbcrlf & "  Folder Does Not Exist. ", "Folder : " & (arrCellData(1)) & " does not exist"
			Else		
				If oFSO.FolderExists(arrCellData(2)) Then 'Checking if folder already exists
					If Right(arrCellData(2),1) <>"\" Then
						arrCellData(2)=arrCellData(2) & "\"
					End If
					oFSO.CopyFolder arrCellData(1), arrCellData(2), True
					Reporter.ReportEvent micDone, "Folder " & arrCellData(1) &" should be copied", " Folder is copied to  Path :" & arrCellData(2)
				Else
					oFSO.CreateFolder(arrCellData(2)) 
				   If Right(arrCellData(2),1) <>"\" Then
						arrCellData(2)=arrCellData(2) & "\"
					End If
					oFSO.CopyFolder arrCellData(1), arrCellData(2), True 'Copying folder from one destination to another  		
					Reporter.ReportEvent micDone, "Destination Folder should be present", " Destination Folder Was Created at Path :" & arrCellData(2)& ", as it did not exist and " & (arrCellData(1)) & "and was Copied"
				End If
			End If
		Case "move"
			If Not oFSO.FolderExists(arrCellData(1)) Then 'Checking if folder already exists
				Reporter.ReportEvent micFail,"Keyword Check at Line no - " & intRowCount &vbcrlf & "  Folder Does Not Exist. ", "Folder " & (arrCellData(1)) & " does not exist"
			Else		
				If oFSO.FolderExists(arrCellData(2)) Then 'Checking if folder already exists
                    DestFolder=(arrCellData(2))& "\" & Foldername
					If oFSO.FolderExists(DestFolder) Then 'Checking if folder already exists
						oFSO.deletefolder(DestFolder) 'Deleting the folder if it already exists
					End If
					oFSO.MoveFolder arrCellData(1), DestFolder 'Moving the desired folder
					Reporter.ReportEvent micDone, "Folder " & arrCellData(1) & " Moved", "Folder moved to destination " & (arrCellData(2))
				Else
					oFSO.CreateFolder(arrCellData(2))
					oFSO.MoveFolder arrCellData(1), arrCellData(2) 'Moving the desired folder 		
					Reporter.ReportEvent micDone, " Destination Folder Was Created As It did not exist",("Folder " & (arrCellData(1)) &" moved")
				End If
			End If
		Case else
			Reporter.ReportEvent micFail, "Invalid Function keyword", "Please check the function keyword mentioned in the row number " &introwcount
	End Select
End Function
'#################################################################################################

'#################################################################################################
'Function name 		: Func_File
'Description        : If the user is working with Files by using FSO 
'Parameters       	: 1.The details to be used while using FSO.                                
'Assumptions     	: NA
'#################################################################################################
'The following function is for 'Function' Keyword
'#################################################################################################
Function  Func_File(pCellData)
	Dim arrFilepath	'Stores the File path
	Dim DestFile	'Stores the Destination File Name
	Dim strFilename	'Stores the File name to be used
	Dim intFileLoc	'Stores the element number of the File Name
	Dim iFSO		'Stores the Created object
	Dim oFile		'Stores the details of Object created
	Dim arrCellData1'Stores the details of the operation to be performed
	Dim intf		'Used for looping
	Dim strMess		'Stores the String which has to be written into a file
	arrCellData1 =split(pCellData,";",-1,1)
	arrFilepath=(split(arrCellData1(1),"\",-1,1))
	intFileloc=UBound(arrFilepath)
	strFilename=arrFilepath(intFileloc)
	Set iFSO = CreateObject("Scripting.FileSystemObject")
	Select Case LCase(arrCellData1(0)) 'select the specific action to be performed on  a file
		Case "create"
			If Not iFSO.FileExists(arrCellData1(1)) Then
				Set oFile = iFSO.CreateTextFile(arrCellData1(1),True)'Creating File
				Reporter.ReportEvent micDone, "File Created", "File: " & (arrCellData1(1)) & " Created"
			Else
				Set oFile = iFSO.GetFile(arrCellData1(1))'Retrieving the File
				Reporter.ReportEvent micDone, "File Already Exists", "File: " & (arrCellData1(1)) & " Exists"
			End If
		Case "delete"	
			If Not iFSO.FileExists(arrCellData1(1)) Then
				Reporter.ReportEvent micFail, "Keyword Check at Line no - " & intRowCount, "File does not Exist at Path: " & (arrCellData1(1))
			Else
				iFSO.DeleteFile(arrCellData1(1))'Deleting File
				Reporter.ReportEvent micDone, "File Deleted", "File: " & (arrCellData1(1)) & " Deleted"
			End If
		Case "copy"
			If Not iFSO.FileExists(arrCellData1(1)) Then
				Reporter.ReportEvent micFail,  "Keyword Check at Line no - " & intRowCount & vbcrlf & "File Does Not Exist" , "File  " & (arrCellData1(1)) & " does not exist"
			Else		
				If iFSO.FolderExists(arrCellData1(2)) Then
					If Right(arrCellData1(2),1) <>"\" Then
						arrCellData1(2)=arrCellData1(2) & "\"
					End If
					iFSO.CopyFile arrCellData1(1), arrCellData1(2) ,true    'Copying File
					Reporter.ReportEvent micDone, "File "&arrCellData1(1) & " Copied", "File copied to " & (arrCellData1(2)) & " location"  	
				Else
					iFSO.CreateFolder(arrCellData1(2))
					iFSO.CopyFile arrCellData1(1), arrCellData1(2), true
					Reporter.ReportEvent micDone, " Destination Folder " &arrCellData1(2)& "  Created As It did not exist",("File " & (arrCellData1(1)) & " Copied")
				End If
			End If
		Case "move"
			If Not iFSO.FileExists(arrCellData1(1)) Then
				Reporter.ReportEvent micFail,"Keyword Check at Line no - " & intRowCount  & vbcrlf & "File Does Not Exist", " File " & (arrCellData1(1)) & " does not exist"
			Else		
				If iFSO.FolderExists(arrCellData1(2)) Then
					If Right(arrCellData1(2),1) <>"\" Then
						arrCellData1(2)=arrCellData1(2) & "\"
					End If
                    DestFile=(arrCellData1(2)) & strFilename
					If iFSO.FileExists(DestFile) Then
						iFSO.deletefile(DestFile)'Deleting the file if it already exists
					End If
					iFSO.MoveFile arrCellData1(1), arrCellData1(2)'Moving file from one place to another
					Reporter.ReportEvent micDone, "File " & arrCellData1(1)& " Moved", "File moved to destination" & (arrCellData1(2))			
				Else
				   iFSO.CreateFolder(arrCellData1(2))
				    arrCellData1(2)=arrCellData1(2) & "\"
					iFSO.MoveFile arrCellData1(1), arrCellData1(2)
					Reporter.ReportEvent micDone, " Destination Folder '" & arrCellData1(2) & "'  was created As It did not exist",("File " & (arrCellData1(1)) & " Copied")
			  End If
			End If
		Case "write"
			Dim arrWrite
			If iFSO.FileExists(arrCellData1(1)) Then
				Set oFile = iFSO.OpenTextFile(arrCellData1(1),2, True)'Opening File to append Text
				intf = 0
				arrWrite = Split(arrCellData1(2),":",-1,1)
				For intf = 0 to  Ubound(arrWrite)
				   If (Instr(1,arrWrite(intf),"#") = 1) Then
					  arrWrite(intf) = GetValue(arrWrite(intf))
				   End If
				 strMess = strMess & " " & arrWrite(intf)
				Next 
				oFile.Write strMess
				Reporter.ReportEvent micDone, "Text Written in File", "Text: '" & strMess & "'  Written"
			Else
				Reporter.ReportEvent micFail,  "File Does Not Exists", "File " & strFilename & " Does Not Exist"
			End If
		Case "read"
			If iFSO.FileExists(arrCellData1(1)) Then
				Set oFile = iFSO.OpenTextFile(arrCellData1(1),1)
				Environment.Value(arrCellData1(2))=oFile.ReadAll
				Reporter.ReportEvent micDone, "Read File Operation Done", ("File " & (arrCellData1(1)) & " read & it contains text " & Environment.Value(arrCellData1(2)))
			Else
				Reporter.ReportEvent micFail,  "File Does Not Exists",("Could not read file as " & strFilename & " Does Not Exist")
			End If
		Case "append"
			If iFSO.FileExists(arrCellData1(1)) Then
				Set oFile = iFSO.OpenTextFile(arrCellData1(1),8, True)
				arrCellData1(2)=GetValue(arrCellData1(2))
				oFile.Write arrCellData1(2)
				Reporter.ReportEvent micDone, "Text Appended to File", "Text: '" & (arrCellData1(2)) & "' appended"
			Else
				Reporter.ReportEvent micFail,  "File Does Not Exists", "File " & strFilename & " Does Not Exist"
			End If
	End Select
End Function
'#################################################################################################

'#################################################################################################
'Function name 		: Func_ExportXML
'Description        : If the user wants to export data and store it in a XML format
'Parameters       	: 1.The details to be used while exporting in XML format.                                
'					  				2.The Path where the XML has to be stored
'Assumptions     	: NA
'#################################################################################################
'The following function is for 'Function' Keyword
'#################################################################################################
Function Func_ExportXML(sPath,strDetails1)
    Dim into			'Used for looping
	Dim oRoot			'Stores the Root Element for the Object
	Dim arrDocSplit		'Stores the elements to be exported to XML and the document name
	Dim strDocName		'Stores the Document name
	Dim arrElementSplit	'Stores the variable and their tag names to be exported to XML
	Dim arrElementName	'Stores the current variable and its tag names to be exported to XML
	Dim oDoc	   		'Stores the XML Object
	arrDocSplit = Split(strDetails1,";",2,1)
	strDocName = arrDocSplit(0)
	arrElementSplit = Split(arrDocSplit(1),"::",-1,1)'Splitting the different values to be stored in the XML
	Set oDoc = Nothing                           ' Before exporting data to file (first time/overwrite) flush the object.
	Set oDoc = XMLUtil.CreateXML()               ' Object to a XML file. 
	oDoc.CreateDocument strDocName          	     ' Name of the file.
	Set oRoot = oDoc.GetRootElement()
	For into = 0 to Ubound(arrElementSplit)
		arrElementName = Split(arrElementSplit(into),":",-1,1)
		arrElementName(0) = GetValue(arrElementName(0))
		arrElementName(1) = GetValue(arrElementName(1))
		oRoot.AddChildElementByName arrElementName(0),arrElementName(1)' Adding data
	Next
	oDoc.SaveFile sPath      					' Saves the file at a particular path.
	Set oRoot = Nothing
	Set oDoc = Nothing 
End Function
'#################################################################################################

'#################################################################################################
'Function name 		: Func_DeleteXML
'Description        : If the user wants to delete a XML file.
'Parameters       	: 1.The path in which the XML File is present.                                
'Assumptions     	: NA
'#################################################################################################
'The following function is for 'Function' Keyword
'#################################################################################################
Function Func_Deletexml(sPath)
	Dim dFileObj	'Stores the XML File to be used
	Dim dFSO	  	'Stores the XML Object
	Set dFSO = Nothing									  ' Before exporting data to file (first time/overwrite) flush the object.
	Set dFSO =  CreateObject("Scripting.FileSystemObject")' Sets an File System Object for the XML file.
	If dFSO.FileExists(sPath) Then
		Set dFileObj = dFSO.GetFile (sPath)
		dFileObj.Delete(True)	   						   ' Verification of the  Deletion of  the xml file.
			Reporter.ReportEvent micPass, "To delete the Existing Xml File", "The Existing XML is deleted, which is as expected"
			Else 
			Reporter.ReportEvent micFail, "To delete the Existing Xml File", "The Xml file does not exists"
   	End If
	Set dFSO = Nothing
End Function
'#################################################################################################

'#################################################################################################
'Function name 		: Func_Report
'Description        : This function provides you the customized report with specified user inputs 
'								through the keyword.
'Parameters       	: None
'Assumptions    	: NA
' Sample Call 		: Func_Report()
'#################################################################################################
'The following function is for Report  Keyword 
'#################################################################################################
Function Func_Report()
   Dim reportobj	'reportobj variable holds the input of the report keyword mention in third Column
	Dim reportcon	'reportcon variable holds the status of the report ex.pass,fail etc. 
	Dim reportcon1	'reportcon1 variable holds actual message of the report.  
	Dim reporter0	'reporter0 variable holds expected message of the report.
	Dim expmess		'stores the concatenated expected message
	Dim actmess		'stores the concatenated actual message
	Dim reporter1	'stores the split  value of reportcon
		reportcon = Split(Rep_Value,";",2,1)  
		reportcon1 = Split(reportcon(1),"::",-1,1) 
		reporter0 = Split(reportcon1(0),":",-1,1)  
		inti = 0
		For inti = 0 to  Ubound(reporter0)
			expmess = expmess & " " & reporter0(inti)
		Next
		inti = 0
		reporter1 = Split(reportcon1(1),":",-1,1)
		For inti = 0 to  Ubound(reporter1)
			actmess = actmess & " " & reporter1(inti)
		Next 
		Select Case Lcase(Trim(reportcon(0))) 'To write the status of the acttion in a report
			 Case "pass"
				Reporter.ReportEvent micPass,expmess,actmess
					If htmlreport = "1" Then
					Call Update_log(MAIN_FOLDER, g_sFileName,"pass")	
				End If
			Case "fail"
				Reporter.ReportEvent micFail,expmess,actmess
				If htmlreport = "1" Then
					Call Update_log(MAIN_FOLDER, g_sFileName,"fail")	
				End If
			Case "done"
				Reporter.ReportEvent micDone,expmess,actmess
			Case "warning"
				Reporter.ReportEvent micWarning,expmess,actmess
			Case Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & reportcon(0) & "'  not supported.Please verify Keyword entered."
				If htmlreport = "1" Then
					Call Update_log(MAIN_FOLDER, g_sFileName,"error")	
				End If
		End Select
End Function
'#################################################################################################

'#################################################################################################
'Function name 	: Func_Condition
'Description    : This function is used to evaluate the expression according to the inputs given in 
'				keyword script.
'Parameters     : intRowCount is a Row number of the keyword script 
'Assumptions    : NA
'#################################################################################################
'The following function is for Condition  Keyword 
'#################################################################################################
Function Func_Condition(ByRef intRowCount)
	Dim iFlag			'Used to set the flag
	Dim cndSplit		'Stores the value of fourth column of local sheet
	Dim startRow		'Contains the start row for the condition
	Dim endRow			'Contains the end row for the condition
    Dim var1						'First element to be evaluated
    Dim var2							'Second element to be evaluated
	iFlag = 2
	cndSplit = Split(objPerform,";",2,1)
	startRow = CInt(cndSplit (0))
	endRow = CInt(cndSplit (1))
	For  inti =0 to Ubound(arrAction)
        If (Lcase(Cstr(arrAction(inti))) ="true") Then
			 arrAction(inti) ="True"
		ElseIf Lcase((Cstr(arrAction(inti))) = "false") Then
			 arrAction(inti) = "False"
		 End If
    		If ( Not(IsDate(arrAction(inti))) and Len(arrAction(inti)) <= 4 and Lcase(arrAction(inti)) <> "blank")Then
			If (Instr(1,arrAction(inti),"#") =0 and arrAction(inti) <> "" ) Then
				If (Asc(arrAction(inti)) <58) Then
				If (Asc(arrAction(inti)) <58) Then
						arrAction(inti) = Trim(arrAction(inti))
						arrAction(inti)=(Cint(arrAction(inti)))
					End If
				End If
			End If
		ElseIf  Lcase(arrAction(inti)) = "blank" Then
			arrAction(inti) = ""
		End If	
	Next
	var1  = arrAction(0)
	var2 =  arrAction(2)
	Select Case LCase(arrAction(1)) 'to select the conditions
		Case "equals"
			If (Eval("var1 = var2")) Then
				iFlag = 0
			End If
		Case "not"
			If NOT(Eval("var1 = var2")) Then
				iFlag = 0
			End If
		Case "greaterthan"
			If (Eval("var1 > var2")) Then
				iFlag = 0
			End If
		Case "lessthan"
	  		If Eval("var1 < var2") Then
				iFlag = 0
			End If
		Case Else
			Reporter.Reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrAction(1) & "'  not supported. Please verify Keyword entered."
	End Select
	If iFlag = 0 Then
	    intRowCount = startRow -1  
	Else
		intRowCount = endRow
	End If
End Function	
'#################################################################################################

'#################################################################################################
'Function name 	   : Func_presskey
'Description       : This function is used for sending keyboard combinations
'Parameters        : NA
'Assumptions       : NA
'#################################################################################################
'The following function is used for 'Presskey' keyword
'#################################################################################################
Function Func_presskey()		' Send Hot Keys Function
	Dim WshShell  'Object created for Shell Scripting
	Set WshShell = CreateObject("WScript.Shell")
	Select Case LCase(arrAction(0)) 'To select the keyboard actions
		Case "enter"
			WshShell.SendKeys "{ENTER}"
		Case "f1"
			WshShell.SendKeys "{F1}"
		Case "f2"
			WshShell.SendKeys "{F2}"
		Case "f3"
			WshShell.SendKeys "{F3}"
		Case "f4"
			WshShell.SendKeys "{F4}"
		Case "f5"
			WshShell.SendKeys "{F5}"
		Case "f6"
			WshShell.SendKeys "{F6}"
		Case "f7"
			WshShell.SendKeys "{F7}"
		Case "f8"
			WshShell.SendKeys "{F8}"
		Case "f9"
			WshShell.SendKeys "{F9}"
		Case "f10"
			WshShell.SendKeys "{F10}"
		Case "f11"
			WshShell.SendKeys "{F11}"
		Case "f12"
			WshShell.SendKeys "{F12}"
		Case "escape"
			WshShell.SendKeys "{ESCAPE}"
		Case "delete"
			WshShell.SendKeys "{DEL}"
		Case "end"
			WshShell.SendKeys "{END}"
		Case "alt+f4"
			WshShell.SendKeys "%{F4}"
		Case "ctrl+s"
			WshShell.SendKeys "^{s}"
		Case "ctrl+p"
			WshShell.SendKeys "^{p}"
		Case Else
			sPresskey=arrAction(0)
			WshShell.SendKeys sPresskey	
	End Select
	Set WshShell=Nothing
End Function
'#####################################################################################################

'####################################################################################################
'Function name 		: Func_StringOperations
'Description        : This function is used for all string operations.
'Parameters       	: The Action keyword in the Second Column of local sheet goes as a parameter
'Assumptions     	: None
'####################################################################################################
'The following function is used for all types of string operations.
'####################################################################################################
Function Func_StringOperations(strCriteria)
   Dim arrSplit 	 'Stores the elements from 3rd Column of the datatable into an array after splitting with ";" delimiter
   Dim strMainString 'Stores the main string 
   Dim strSubString  'stores the sub string
   Dim intLen        'Stores the length of the array "arrSplit"
   Dim ReturnVal     'stores the Return value for example the position in case of string search operation
   Select Case strCriteria 'Case statement for the string operations
 	Case "strsearch"
		ReturnVal = ""
        ReturnVal = Instr(1,arrAction(0),arrAction(1)) 'Searching for strSubString in strMainString and storing the position in ReturnVal
	Case "strconcat"
		ReturnVal = ""
		arrSplit = Split(Rep_Value,";")
		intLen = Ubound(arrSplit)
		For inti=0 To intLen
			ReturnVal = ReturnVal & arrSplit(inti) 'concatenating the specified strings and storing in ReturnVal
		Next 	
	Case "strreplace"
		Dim strString 'Stores the string that will replace strSubstring within strMainString.
		ReturnVal = ""
		arrSplit = Split(Rep_Value,";")
		ReturnVal = Replace (arrSplit(0),arrSplit(1),arrSplit(2)) 'Replacing the strSubString with strString in strMainString
	Case "strreverse"
		 ReturnVal = ""
         ReturnVal=strReverse(Rep_Value)'reverse the specified strings and store in ReturnVal
   Case else
	   Reporter.ReportEvent micFail, "Invalid string operation", "Invalid string operation is mentioned in the keyword test step no " &introwcount & ". Plese check the keyword test script."
   End Select
    Func_StringOperations = ReturnVal		'Stores the end result into the Variable specified in the 4th Column
End Function
'########################################################################################################

'########################################################################################################
''Function name 		: Func_gfRegExpTest
''Description        : This function conducts a Regular Expression test.
''Parameters       	: The pattern string to be searched for in the main string.
''Return Value 		: True or False
''Sample call		: Call Func_gfRegExpTest("TestStri.#", "TestString")
''Assumptions     	: NA
''####################################################################################################
''The following function is used for working with regular expressions
''####################################################################################################
Function Func_gfRegExpTest(strPattern, strString)
	Dim objRegEx            ' Create variable.
	Set objRegEx = New RegExp         		' Create regular expression.
	objRegEx.Pattern = strPattern     				' Set pattern.
	objRegEx.IgnoreCase = False      				' Set case sensitivity.
	Func_gfRegExpTest = objRegEx.Test(strString) ' Execute the search test.  This will return a True/False
End Function
''####################################################################################################

'####################################################################################################
''Function name 		: Func_gfQuery
''Description        : If the user requires to query a Database then this function can be used.
''Parameters       	: strSQL  - The Query that needs to be executed                            
''Assumptions     	: The connection string is specified and a connection is established with the database.
''Sample Call		: Func_gfQuery(Select employee_id from employee)
''####################################################################################################
''The following function is used for Performing Database(SQL) Operations
''####################################################################################################
Function Func_gfQuery(strSQL)
   	 For inti = 0 to (Func_RegExpMatch ("##\w###",strSQL,aPosition,aMatch) - 1)
		strReplace = "'" & Environment.Value(Replace(aMatch(inti),"##","",1,-1,1)) & "'"
		strSQL = Func_gfRegExpReplace(aMatch(inti), strSQL, strReplace)
	Next
	'Create the DB Object
	Set dbConn = CreateObject("ADODB.Connection")
	dbConn.Open Environment.Value("connectionString")
	'Execute the query
   	 Set dbRs = dbConn.Execute(strSQL)
     'This condition is used for error handling
	 If dbRs.State Then
		If (dbRs.EOF Or dbRs.BOF) Then
			Reporter.ReportEvent micDone ,"Keyword Check at Line No: " & intRowCount  ,"The Query: " & strSQL & " , should return a value. The Query did not return a value."
			Func_gfQuery = "<dne>"
		ElseIf isNull(dbRs.Fields.Item(0)) Then
            Func_gfQuery = Empty
		Else
			Func_gfQuery = dbRs.Fields.Item(0) 
		End If
	End If
    'Close the database connection
	dbConn.Close
	Set dbConn = Nothing
End Function
'#############################################################################################################

'#############################################################################################################
'Function name 		: Func_RegExpMatch
'Description        : Executes a regular expression search against a specified string
'Parameters       	: The pattern string to be searched for in the main string
'				 	 					The main string 
'Return Value		: Returns a Match collection when a regular expression search is performed. Reference parameters are used to return the start position(aIndex) and value(aValue)
'Assumptions     	: NA
'#############################################################################################################
'The following function is used to perform a regular expression search
''#############################################################################################################
Function Func_RegExpMatch(strPattern,strString, aIndex(), aValue())
	Dim regEx, Match, Matches   			' Create variable.
	Set regEx = New RegExp   		  		' Create regular expression.
	regEx.Pattern = strPattern  	   			  ' Set pattern.
	regEx.IgnoreCase = True   			    	' Set case insensitivity.
	regEx.Global = True  					             ' Set global applicability.
	Set Matches = regEx.Execute(strString)  ' Execute search.
	'Resize the arrays
	ReDim aIndex(Matches.Count - 1)
	ReDim aValue(Matches.Count - 1)
	inti = 0
	For Each Match in Matches  	 		' Iterate Matches collection.
       		aIndex(inti) = Match.FirstIndex
		aValue(inti) = Match.Value
		inti = inti + 1
	Next
	Func_RegExpMatch = Matches.Count
End Function
'########################################################################################################

'########################################################################################################
'Function name 		: Func_gfRegExpReplace
'Description        : Replaces text found in a regular expression search
'Parameters       	: The pattern string to be searched for and replaced
'			  	    The main string in which the pattern string needs to be replaced
'			   		The string which replaces the pattern string found in main string
'Return Value		: Returns replaced text
'Assumptions     	: NA
'#######################################################################################################
'The following function is used to replace a string using regular expression search
'#######################################################################################################
Function Func_gfRegExpReplace(strPattern, strFind, strReplace)
	Dim regEx               				'Create variables.
	Set regEx = New RegExp     				'Create regular expression.
	regEx.Pattern = strPattern		  		'Set pattern.
	regEx.IgnoreCase = True      			'Make case insensitive.
	Func_gfRegExpReplace = regEx.Replace(strFind, strReplace)   ' Make replacement.
End Function
'##########################################################################################################

'#####################################################################################################
'Function name 		: Func_loop
'Description           : Repeats a group of statements a specified number of times
'Parameters          : NA
'Assumptions     	 : If the number of times to be looped is not specified, by default this number is taken as the number of active rows 
'					    	in Action1 sheet of Datatable.
'#####################################################################################################
'The following function is used for Loop Keyword
'#####################################################################################################
 Function Func_loop()
'	Dim arrloopData 'Stores the start row and end row values
	Dim intcntr					'Stores the loop count
	Dim Counter					'Stores the Count value
	Dim endRow1					'Stores the End row for looping
	Dim loopRowCount			'Stores the current loop count
	Dim StrtRow1							'stores the arrAction(0) value
	StrtRow1 = arrAction(0)
	endRow1 = arrAction(1)
	If isArray(arrkeyIndex) Then
		intcntr= arrkeyIndex(0)
	Else
		intcntr= Datatable.GlobalSheet.GetRowCount   
	End If
	If isnumeric(StrtRow1) and isnumeric(endRow1) and isnumeric(intcntr) Then
		Counter=Cint(intcntr)
		intDataCounter=1
			Do While intDataCounter <= Counter 'Checking if End Row is greater than Start Row
					intRowCount = Cint(StrtRow1)
					endRow1 = Cint(endRow1)
					Datatable.LocalSheet.SetCurrentRow(intDataCounter)
					For loopRowCount = intRowCount to endRow1
						While intRowCount<=endRow1
							Datatable.LocalSheet.SetCurrentRow(intRowCount)
							wait 1
							If (LCase(DataTable.Value(1,dtLocalSheet)) = "r") Then
								Call Keyword_Call()
							End If
							If initial = "condition" Then
								intRowCount = intRowCount + 1	'intRowCount changes depending upon the number of lines in condition
							Else
								intRowCount = intRowCount + 1 	'intRowCount changes depending upon the current row and increments by 1
							End If
						Wend
					Exit For
					Next
					intDataCounter = intDataCounter + 1
				Loop
				intRowCount = endRow1
				intDataCounter = 1
	else
		Reporter.ReportEvent micFail, "Loop Exception occurred", "Please provide numerical values for the loop keyword. Please check the keyword test step " & intRowCount & "."
	End if
 End Function

'############################################################################################################

'############################################################################################################
'Function name 		: Func_Convert
'Description        : This function is used to convert from one data type to another.
'Parameters       	: The variable which is to be converted and the variable name in which the converted  type is to be stored
'Return Value		: This function returns the converted variable.
'Assumptions     	: NA
'############################################################################################################
'The following function is used for convert Keyword
'############################################################################################################
Function Func_Convert()
	Dim Conv_val 'Stores the converted value
	Select Case LCase(arrkeyIndex(0)) 'Convert one data type to another
		Case "date"
		 Select Case LCase(arrkeyIndex(2)) 'converting the date format
			Case "0"
				Conv_val = FormatDateTime(arrAction(0),0)
			Case "1"	
				Conv_val = FormatDateTime(arrAction(0),1)
			Case "2"
				Conv_val = FormatDateTime(arrAction(0),2)
			Case "3"
				Conv_val = FormatDateTime(arrAction(0),3)
			Case "4"   
				Conv_val = FormatDateTime(arrAction(0),4)
			Case Else
				Reporter.ReportEvent micFail,"Incorrect Date/Time Format","Check The Format specified in the row number " & introwCount
			End Select
	Case "roundto"
		If Ubound(arrkeyIndex) = 2 Then
			Conv_val = Round(arrAction(0),Cint(arrkeyIndex(2)))
		else
			Conv_val = Round(arrAction(0))
		End If		
	Case "lcase"
		Conv_val = LCase(arrAction(0))
	Case "ucase"
		Conv_val = UCase(arrAction(0))
	Case "cstr"
		Conv_val = Cstr(arrAction(0))
	Case "ascii"
		Conv_val = Asc(arrAction(0))
	Case "trim"
		 Conv_val = Trim(arrAction(0))
	Case "len"
		Conv_val = len(arrAction(0))
	Case Else
		Reporter.ReportEvent micFail,"Invalid convert operation","Convert operation mentioned in the keyword test step no: " & introwCount & " is not supported."
   End Select
   Func_Convert = Conv_val
End Function
'##########################################################################################################

'##########################################################################################################
'Function name 	: Func_ImportData
'Description    : It is used to import test data at runtime
'Parameters     : The file name and the path where it is stored
'Assumptions    : The required file is present. The required file is an excel sheet. 
'##########################################################################################################
'The following function is used for keyword 'importdata'.
'##########################################################################################################
Function Func_ImportData(strTestCase)
    Dim strPath    			 '	stores the path with .xls 
    Dim strDataPath 		'locates and stores the full path 
	Dim strSheetName		'stores the sheet name to be imported
    strTestCase = CStr(Trim(strTestCase))
    If InStr(1,strTestCase,".xls") = 0 Then	   'check if .xls is given in the path provided.
       strPath     = strTestCase & ".xls"		   'adding .xls to the Path
    Else
       strPath     = strTestCase
    End If
	If isArray(arrKeyValue) = True Then
		strSheetName = arrKeyValue(0)
		If (Asc(strSheetName) <58) Then
					If (Asc(strSheetName)>47) Then
						strSheetName = Cint(Trim(strSheetName))
					End If
				End If
	Else
		strSheetName = 1
	End If
    Wait 2
    If InStr(1,strTestCase,":\") <> 0 Then
        strDataPath = strPath	           'storing the full path in strDataPath
    Else
        strDataPath = Pathfinder.Locate(strPath) 'locating and storing the full path in strDataPath
    End If
    Wait 2
    If strDataPath <> "" Then	'Check if the sheet is present in the given path
        DataTable.ImportSheet strDataPath,strSheetName, 1  'If present import the data into Action1 sheet
    Else
		Reporter.ReportEvent micFail,"Incorrect Input","The file" & " " & strPath & " " & "is not present in the attachment for the current test. Please check the value of the row " & intRowCount
    End If
End Function
'##########################################################################################################

'##########################################################################################################
'Function name 		: Func_Eloop
'Description        : It is used fto identify the end row when a loop keyword is called in the script
'Parameters       	: None
'Assumptions     	: None
'###########################################################################################################
Function Func_Eloop()
		Dim lpno 'stores the upperbound value of loopcnt
		lpno = Ubound(Loopcnt)
		Loopind(lpno)= Loopind(lpno) + 1
		intdatacounter = Loopind(lpno)+1
		If (Loopind(lpno)=CInt(Loopcnt(lpno))) Then
			ReDim preserve Loopind(lpno-1)
			ReDim preserve Loopcnt(lpno-1)
			ReDim preserve LpStrtRow(lpno-1)
			Func_Eloop = 0
		else
		    Func_Eloop = LpStrtRow(lpno)
		End If
End Function
'#####################################################################################################

'#####################################################################################################
'Function name 		: Func_goto
'Description        :  This function to the parameter  label name and start execute from that row in datatable 
'Parameters       	: Label name
'Assumptions     	: None
'#####################################################################################################
Function Func_goto(Lblname)
	Dim g  			'used for looping
	Dim g_flag	 'initialization variable
	g_flag = 0
	For g=1 to Datatable.LocalSheet.GetRowCount
		Datatable.SetCurrentRow(g)
		If ((Lcase(Datatable.Value(1,dtLocalSheet)) = "r" ))Then
			If ((Lcase(Datatable.Value(2,dtLocalSheet)) = "label" )) Then
					If ((Lcase(Datatable.Value(3,dtLocalSheet)) = lcase(Lblname) )) Then
						g_flag = 1
						ecflag = cflag
						EloopNo = LoopNo
						Func_goto = g
						Exit function
                    End If
			End If
         End If
	Next
If g_flag = 0 Then
	Reporter.ReportEvent micFail, "Label Exception Occurred", "Could not find the label " + Lblname + " in the keyword test script"
	Func_goto = 0
End If
End Function
'#########################################################################################################

'########################################################################################################
'Function name 		: Func_condition1
'Description        : 
'Parameters       	: 
'Assumptions     	: NA
'########################################################################################################

'########################################################################################################
Function Func_Condition1( )
   Dim condflag  	'initialization variable
   Dim val1 				'stores the condition value 1
   Dim oper 					'variable used in select case for choosing the type of operations
   Dim val2 						'stores the condition value 2
   Dim iFlag						 'stores the flag value
   iFlag = 1
	condflag = 0
	For  inti =0 to Ubound(arrAction)
		If (Lcase(Cstr(arrAction(inti))) ="true") Then
			 arrAction(inti) ="True"
		ElseIf Lcase((Cstr(arrAction(inti))) = "false") Then
			 arrAction(inti) = "False"
		 End If
		If ( Not(IsDate(arrAction(inti))) and Len(arrAction(inti)) <= 4 and Lcase(arrAction(inti)) <> "blank")Then
			If (Instr(1,arrAction(inti),"#") =0 and arrAction(inti) <> "" ) Then
				If (Asc(arrAction(inti)) <58) Then
					If (Asc(arrAction(inti))>47) Then
						arrAction(inti) = Trim(arrAction(inti))
						arrAction(inti)=(Cint(arrAction(inti)))
					End If
				End If
			End If
		ElseIf  Lcase(arrAction(inti)) = "blank" Then
			arrAction(inti) = ""
		End If	
	Next
   val1 = arrAction(0)
   oper = arrAction(1)
   val2 = arrAction(2)
Select Case LCase(oper) 'select statement for checking the conditions
		Case "equals"
			If (val1 = val2) Then
				iFlag = 0
			End If
		Case "not"
			If val1 <> val2 Then
				iFlag = 0
			End If
		Case "greaterthan"
			If (isnumeric(val1) and isnumeric(val2))  Then
				If Cint(val1) > Cint(val2) Then
					iFlag = 0
				End If
			else
				Reporter.ReportEvent micFail, "Condition exception occurred", "Please provide numeral values for the condition keyword used in the step " &introwcount & "."
			end if
		Case "lessthan"
	  		If (isnumeric(val1) and isnumeric(val2))  Then
				If Cint(val1) < Cint(val2) Then
					iFlag = 0
				End If
			else
				Reporter.ReportEvent micFail, "Condition exception occurred", "Please provide numeral values for the condition keyword used in the step " &introwcount & "."
			end if
		Case Else
			Reporter.Reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrAction(1) & "'  not supported." , "Please verify Keyword entered."
	End Select

If iflag=0 Then
        Func_Condition1 = 0
		Exit function
else
	Dim rowtemp
	Dim ir 'used for looping
	Dim drcnt 'stores the row count
	drcnt = Datatable.LocalSheet.GetRowCount
	rowtemp = intRowCount
	For ir = rowtemp + 1 to drcnt
		Datatable.LocalSheet.SetCurrentRow(ir)
		If Lcase(Datatable.Value(1, dtlocalSheet)) = "r" Then
			If (Lcase(Datatable.Value(2,dtlocalSheet)) = "condition") Then
				condflag = condflag + 1
			End If
			If (Lcase(Datatable.Value(2,dtlocalSheet)) = "endcondition") Then
				If (condflag = 0) Then
					Reporter.Reportevent micDone,  "Condition check performed at row number " & intRowCount &" failed", "Condition check failed"
					Func_Condition1 = ir-1
					Exit function
				End If
				If (condflag <> 0) Then
					condflag = condflag - 1
				End If
			End If
		End If
	Next
End If
End Function
'########################################################################################################

'########################################################################################################
'Function name 	   : Func_CaptureScreenshot 			
'Description       : This function is used for Capturing screenshot
'Parameters        : string value depending upon  on which the file name of the screenshot will differ
'Assumptions       : NA
'#######################################################################################################
'The following function is used for 'screencapture' keyword
'#######################################################################################################
Function Func_CaptureScreenshot(strtype,intRowCount)		
	Dim strFolderPath	'stores the folder path , where to save the screen shot
    Dim oFSO 					 'Object created for Shell Scripting
	Dim strTimestamp		'stores the folder with the time, day and yr of the current system
	Dim strTestName				'variable used to store the test name
	Dim oFolder 						 'object created to access folders
	Dim strTestPath						 'path of the screenshot with the testname 
	Dim strDestination 						'path stores where the screenshot is stored
	Dim strFilepath									'entire path where the screenshot is saved
	Dim intDone  											'used for initialization                                   
	strFolderPath = "C:\O2T"   						 ' master folder path. Change this to NA to skip screenshot feature
	If lcase(strFolderPath) <>  "na" Then
	strTestName = environment("TestName")
	wait(1)
	strTimestamp = cstr(month(now) & "-"  & day(now) & "-" & year(now) &"_" & hour(now) & "-" & minute(now) & "-" & second(now))

    Set oFSO = CreateObject("Scripting.FileSystemObject")'Creating a FSO object

	If Not oFSO.FolderExists(strFolderPath) Then 				'Checking if  the master folder O2TScreenshots exists under C: drive
				Set oFolder = oFSO.CreateFolder(strFolderPath)		'Creating the master folder
				Set oFolder = nothing
	End IF

	strTestPath = cstr(strFolderPath & "/" & strTestName)

	If Not oFSO.FolderExists(strTestPath) Then 							'Checking if a folder for the specific test already exists -TestName
				Set oFolder = oFSO.CreateFolder(strTestPath)		'Creating a folder for the specific test
				Set oFolder = nothing
	End IF  

	If intDone <> 1 Then	
		strDestination = strTestPath & "/" & "Run" '& strTimestamp
		If Not oFSO.FolderExists(strDestination) Then 	
			oFolder = oFSO.CreateFolder(strDestination)  					'create a folder for the specific test run session - Run_timestamp
        	Set oFSO=Nothing
		End If
		destinationpath = strDestination         'to store the folder created for the current run session
	Else
		strDestination = destinationpath						'store the folder path already created
	End if

	If lcase(strType) = "test" Then
		strFilepath = strDestination & "/" & strTestName & "_" & strTimeStamp &", Row No. - "& intRowCount  & ".png" 'Filename = Name of test_timestamp.png		
	Else
		strFilepath  = strDestination & "/" & strTestName & "_CheckpointFailure_Row No. - " & intRowCount & ".png"
	End If
	curParent.CaptureBitmap strFilepath 				'this stores the screenshot onto the local system                                                      
	intDone = 1
	End if
End Function
'################################################################################################################

'################################################################################################################
'Function Name   			 : createlog
'Description      				: To create an execution log in the HTML file
'Input Parameters 		:  path of the log to be stored and value from 4th column  of the datatable
'Output Parameters: None
'#################################################################################################################
Function createlog(objName, objPerform)
	Dim g_objReport			'File Object
	Dim g_objFS						'File System Object
	Dim strTimestamp 			'Time stamp of the current system
	Dim strIntPath   						'stores the path, where the execution log is summarized
	Dim g_iPass_Count 					'initialization variable
	Dim g_iFail_Count  						'initialization variable
	Dim g_iImage_Capture  				'initialization variable
	Dim g_Total_TC								 'initialization variable
	Dim g_Total_Pass  							'initialization variable
	Dim g_Total_Fail 								'initialization variable
	Dim g_Flag											  'initialization variable
	Dim g_Flag1 											'initialization variable
	Dim g_ScriptName								 'stores the test name
	Dim strResultPath  										'Stores the folder structure
	Dim objFS  														'QTP object created to access the file objects
	Dim objMain  													'QTP object
	Dim g_tStart_Time 											'stores the current time of the system
	Dim sfile 																'vairable used to store the log file with execution time stamp

	strTimestamp = cstr(month(now) & "-"  & day(now) & "-" & year(now) &"_" & hour(now) & "-" & minute(now) & "-" & second(now))
	strIntPath = objName
	sfile = objPerform
	If objName = "" Then
        		strIntPath = "C:\O2T\"  
	End If
	g_iPass_Count = 0
	g_iFail_Count = 0
	g_iImage_Capture = 1
	g_Total_TC = 0
	g_Total_Pass = 0
	g_Total_Fail = 0
	g_Flag = 0
	g_Flag1=0
	g_ScriptName=Environment.Value("TestName")
	MAIN_FOLDER=strIntPath
	Set objFS = CreateObject("Scripting.FileSystemObject")
	If Not objFS.FolderExists(strIntPath) Then
		Set objMain = objFS.CreateFolder(strIntPath)
	End If
	If LCase(sfile) = "yes" Then
		g_sFileName = Environment.Value("TestName")&strTimestamp&".html" 
	Else
		g_sFileName = Environment.Value("TestName")&".html" 
	End If
	Set g_objFS = CreateObject("Scripting.FileSystemObject")
	Set g_objReport = g_objFS.OpenTextFile(MAIN_FOLDER&g_sFileName, 2, True)
	g_objReport.Write "<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"
	g_objReport.Write "<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"
	g_objReport.Write "<TR><TD BGCOLOR=#66699 WIDTH=27%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Scenario Name:</B></FONT></TD><TD COLSPAN=6 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>" & Environment.Value("TestName") & "</B></FONT></TD></TR>"
	g_objReport.Write "<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>"
	g_objReport.Write "<TR COLS=6><TD BGCOLOR=#FFCC99 WIDTH=3%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Row</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=15%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Keyword</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Object</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Action</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Execution Time</B></FONT></TD><TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Status</B></FONT></TD></TR>"
    g_objReport.Close
	Set g_objFS = Nothing 'releasing the buffer value
	Set g_objReport = Nothing 'Releasing the buffer value
	g_tStart_Time = Now() 'Current system time
	Reporter.ReportEvent micDone, "To create an execution log in the HTML file","Execution Log is created in an HTML File"
End Function
'##############################################################################################################

'##############################################################################################################
'Function Name    : Update_log
'Description      : To update the status of the created execution log in the HTML file
'Input Parameters : status
'Output Parameters: 
'##############################################################################################################
Function Update_log(MAIN_FOLDER, g_sFileName, status)
	Dim  strTime	 'stores the current system time
	Dim  fso 				'QTP object created to access files
	Dim fso1				 'Object created for accessin files
	Dim stat 					 'stores the report status
	Dim stat1 						'stores the report status
	strTime = cstr(hour(now) & "-" & minute(now) & "-" & second(now))
	Set fso = createObject("scripting.filesystemobject")
	Set fso1 = fso.OpenTextFile(MAIN_FOLDER&g_sFileName ,8,true)
	Select Case lcase(status) 'To write the status of the execution
		case "executed"
			If datatable.Value(4,dtLocalSheet)<>"" Then
				fso1.Write "<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>" & introwcount & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>" & datatable.Value(2,dtLocalSheet) & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"&datatable.Value(3,dtLocalSheet)&"</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"&datatable.Value(4,dtLocalSheet)&"</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"&strTime&"</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"&status&"</FONT></TD></TR>"
			else
				fso1.Write "<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>" & introwcount & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>" & datatable.Value(2,dtLocalSheet) & "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"&datatable.Value(3,dtLocalSheet)&"</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2><DIV ALIGN = CENTER>"&"----------"&"</DIV></FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"&strTime&"</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"&status&"</FONT></TD></TR>"
			End If
		Case "error"
			fso1.Write "<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred @ Line" & introwcount & "Description:"& Err.number &"-"&Err.description&"</div></th></FONT></TR>"
		Case "pass"
			fso1.Write "<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=GREEN><div align=left>P </FONT><FONT FACE=VERDANA SIZE=2>Reporter Passed @ Line" & introwcount & " Reporter Statement:"&DataTable.Value(3,dtLocalSheet)&"</div></th></FONT></TR>"
		Case "fail"
			fso1.Write "<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=RED><div align=left>O </FONT><FONT FACE=VERDANA SIZE=2>Reporter Failed @ Line" & introwcount & " Reporter Statement:"&DataTable.Value(3,dtLocalSheet)&"</div></th></FONT></TR>"
		Case "checkfail"
			fso1.Write "<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=RED><div align=left>O </FONT><FONT FACE=VERDANA SIZE=2>Check Failed @ Line" & introwcount &"</div></th></FONT></TR>"
		Case "checkpass"
			fso1.Write "<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=GREEN><div align=left>P </FONT><FONT FACE=VERDANA SIZE=2>Check Passed @ Line" & introwcount &"</th></FONT></div></TR>"
		Case "rastart"
			fso1.Write "<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=BLUE><div align=left>C </FONT><FONT FACE=VERDANA SIZE=2>Entering Reusable Action : "&Environment.Value("ReusableActionName")&"</div></th></FONT></TR>"
		Case "raend"
			fso1.Write "<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=BLUE><div align=left>XB </FONT><FONT FACE=VERDANA SIZE=2>Exiting Reusable Action : "&Environment.Value("ReusableActionName")&"</div></th></FONT></TR>"
		Case "finish"
			stat=Reporter.RunStatus
			Select Case stat 'Writing the status of the execution
				Case micFail
					stat1="Fail"
   				Case else
					stat1="Pass"				
			End Select
			If stat1 = "Pass" Then
				fso1.Write "<TR COLS=6><th colspan = 5 BGCOLOR=#EEEEEE><FONT FACE=VERDANA SIZE=2>Execution Completed for: "&Environment.Value("TestName")&"</th></FONT><th BGCOLOR=#EEEEEE><FONT FACE=VERDANA SIZE=2>Status : </FONT><FONT FACE='WINGDINGS 2' SIZE=2 COLOR=GREEN>P</FONT><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>PASS</FONT></TR>"
				else
				fso1.Write "<TR COLS=6><th colspan = 5 BGCOLOR=#EEEEEE><FONT FACE=VERDANA SIZE=2>Execution Completed for: "&Environment.Value("TestName")&"</th></FONT><th BGCOLOR=#EEEEEE><FONT FACE=VERDANA SIZE=2>Status : </FONT><FONT FACE='WINGDINGS 2' SIZE=2 COLOR=RED>O</FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>FAIL</FONT></TR>"
			End If
	End Select
	fso1.Close
	Set fso1=nothing
	Set fso=nothing
End Function
'####################################################################################################################

'####################################################################################################################
'Function name 	    : Func_Wait
'Description        : This function is used for synchronization  with the application
'Parameters       	: The 'Object type' and the 'action being performed is passed as parameters. 
'Assumptions     	: None
'####################################################################################################################
'The following function is used internally.
'####################################################################################################################
Function Func_Wait(arrObjchk,arrKeyValue,initial)
	On error resume next
	If Lcase(Trim(initial)) = "perform" or Lcase(Trim(initial)) = "context" Then
		   If(Ubound(arrKeyValue) >= 0)  Then
				If (LCase((arrKeyValue(0)) <> "exist" ) and (LCase(arrKeyValue(0)) <> "visible" )) Then
					curParent.WaitProperty "visible",True,20
				End If
				If (LCase((arrKeyValue(0)) <> "enabled") And LCase((arrKeyValue(0)) <>"exist") and LCase((arrKeyValue(0)) <> "visible" )) Then
					If LCase((arrObjchk(0)) = "button" Or LCase(arrObjchk(0)) = "checkbox" Or LCase(arrObjchk(0)) = "textbox" Or LCase(arrObjchk(0)) ="radiobutton" Or LCase(arrObjchk(0)) = "tablebutton" Or LCase(arrObjchk(0)) = "tablecheckbox" Or LCase(arrObjchk(0)) = "tabletextbox" Or LCase(arrObjchk(0)) ="tableradiobutton") Then
						object.WaitProperty "visible",True,20
					ElseIf LCase(arrObjchk(0)) = "table" Or LCase(arrObjchk(0)) = "combobox" Or LCase(arrObjchk(0)) = "listbox" Or LCase(arrObjchk(0))= "element" Or LCase(arrObjchk(0)) = "link"  Or LCase(arrObjchk(0)) = "tablecombobox" Or LCase(arrObjchk(0))= "tableelement" Or LCase(arrObjchk(0)) = "tablelink"   Then
						object.WaitProperty "visible",True,20
					End If
				End If
			End If
		   End If 
End Function
'#####################################################################################################################

'#####################################################################################################################
'Function name 	    :  Func_TechInitialize
'Description        :  This function associate the execution to start with object specific technology  incorporating the Framework functions 
'Parameters       	: None
'Assumptions     	: None
'######################################################################################################################
Public Function Func_TechInitialize()
   On error resume next
   vTech = 0
   ReDim arrTechList(8)
   arrTechList(0) = "web"
   arrTechList(1) = "win"
   arrTechList(2) = "java"
   arrTechList(3) = "dotnet"
   arrTechList(4) = "mf"
   arrTechList(5) = "orac"
   arrTechList(6)="Pb"
   arrTechList(7) = "flex"
   arrTechList(8) = "sap"
   Redim arrTech(0)
   Dim k
   Dim j
	For k=0 to ubound(arrTechList)
		Err.clear
		Execute "Call Keyword_" & arrTechList(k) & "()"
		If err.number = 450 Then
			arrTech(j) = arrTechList(k)
			j=j+1	
			ReDim preserve arrTech(j)
		End If
	Next
   ReDim preserve arrTech(j-1)
   Err.clear
End Function
'###############################################################################################################

'##############################################################################################################
'Function name 		: Func_GetXMLValue
'Description           : If the user wants to validate the XML node Value..
'Parameters       	:  1.The path in which the XML File is present. 
'								    2.	The Element Name
'								   3. The Variable	Name for the returning the element value				   
'Assumptions     	: NA
'##############################################################################################################
Function Func_GetXMLValue(sDetails)
Dim oDoc						 'Holds the XML object
Dim blnXmlFlag					'Holds the Flag Value
Dim strArrArg					 		'Holds the XML path,Element Name and the Variable name where the Value to be stored
Dim sPath 									'Holds the Path of the XML file
Dim strElementName					'Holds the Element name to be searched
Dim intChildCnt									 'Holds the Node count
Dim oRootEle 										'Holds the Root element object
Dim oChildEle 											'Holds the Child element object
Dim strActvariable 										'Holds the element Value 
Dim intCnt 															'User for the For counter
Dim introotele
Dim elename															'Element name of the Child 
Dim childcnt																'Count the no. of child elements 
Dim childelename 														'Holds the element name of the child attribute 
Dim intcnt1 																		'used for the increment counter 
Dim childatt 																			' Holds the value of the child attribute 
Dim strElementName1																'Holds the element to be searched
Dim NodeVal
Dim varname 																					'Stores the value of the child element value
blnXmlFlag=False

strArrArg=Split(sDetails,";") 'Splits the XML file path , Element name and the Variable name
sPath=GetValue(strArrArg(0)) 'path of the XML file location
strElementName=GetValue(strArrArg(1))
strElementName1=GetValue(strArrArg(2))
Varname=strArrArg(3) 'Variable name where it stores the value of the element

Set oDoc=XMLUtil.CreateXMLFromFile(sPath) 'Returns the XML File object reference
Set oRootEle= oDoc.GetRootElement
Set oChildEle=oRootEle.ChildElements
intChildCnt=oChildEle.Count 'used to store the no of attributes in one node
For intCnt=1 to intChildCnt
elename=oChildEle.Item(intCnt).ElementName
		If Lcase(elename)=Lcase(strElementName) Then
			childcnt=oChildEle.Item(intCnt).ChildElements.count
        				For intcnt1=1 to childcnt
						childelename=oChildEle.Item(intCnt).ChildElements.Item(intcnt1).ElementName
						If LCase (childelename)=LCase(strElementName1) Then
						childatt=oChildEle.Item(intCnt).ChildElements.Item(intcnt1).Value						
						strArrArg(3)=childatt
                         Environment(Lcase(Varname))=childatt 'stores the value of the variable in an environmental variable
						Reporter.ReportEvent micPass, "Value at Node  '" & childelename &"' of '" &elename &"'", "Value at Node '" &childelename &"' is '" &strArrArg(3) &"' of '" &elename &"'  is as expected."
						If htmlreport = "1" Then
							Call Update_log(MAIN_FOLDER, g_sFileName,"pass")	
						End If
						blnXmlFlag=True 
				Exit for
   						End If
				Next
		End If
Next
			If blnXmlFlag=False Then
				Reporter.ReportEvent micFail , "The Element not  found " ,"The element  does not exists in XMLFile"
			End If
Set oChildEle=Nothing
Set oRootEle=Nothing
Set oDoc=Nothing

End Function
'####################################################################################################################

'####################################################################################################################
'Function name 		: Func_SetXMLValue
'Description        : If the user wants to Set the XML attribute value..
'Parameters       	:  1.The path in which the XML File is present. 
'								    2.	The Element Name
'								   3. The Value to set				   
'Assumptions     	: NA
'###################################################################################################################
'The following function is for 'Function' Keyword
'###################################################################################################################
Public Function Func_SetXMLValue(sDetails)
Dim oDoc					'Holds the XML object
Dim blnXmlFlag				'Holds the Flag Value
Dim strArrArg 						'Holds the XML path,Element Name and the Variable name where the Value to be stored
Dim sPath 									'Holds the Path of the XML file
Dim strElementName					'Holds the Element name to be searched
Dim intChildCnt 'Holds the Node count
Dim oRootEle 'Holds the Root element object
Dim oChildEle											 'Holds the Child element object
Dim strActvariable 											'Holds the element Value 
Dim intCnt 																'User for the For counter
Dim elename																'Element name of the Child 
Dim childcnt																	'Count the no. of child elements 
Dim childelename 															'Holds the element name of the child attribute 
Dim intcnt1																					 'used for the increment counter 
Dim childatt 																					' Holds the value of the child attribute 
Dim strElementName1																		'Holds the element to be searched
blnXmlFlag=False																					'Initialize the flag to False
strArrArg=Split(sDetails,";") 'Splits the XML file path , Element name and the Variable name
sPath=GetValue(strArrArg(0))  ' path of the XML file location
strElementName=GetValue(strArrArg(1))
strElementName1=GetValue(strArrArg(2))
Set oDoc=XMLUtil.CreateXMLFromFile(sPath) 'Returns the XML File object reference
Set oRootEle= oDoc.GetRootElement
Set oChildEle=oRootEle.ChildElements
intChildCnt=oChildEle.Count
For intCnt=1 to intChildCnt
elename=oChildEle.Item(intCnt).ElementName
		If Lcase(elename)=Lcase(strElementName) Then
			childcnt=oChildEle.Item(intCnt).ChildElements.count
        				For intcnt1=1 to childcnt
						childelename=oChildEle.Item(intCnt).ChildElements.Item(intcnt1).ElementName
						If LCase(childelename)=LCase(strElementName1) Then                            					
						oChildEle.Item(intCnt).ChildElements.Item(intcnt1).SetValue(GetValue(strArrArg(3)))
						Reporter.ReportEvent micPass, "Value at Node  " & childelename &" to be set " , "Value at Node " & childelename &" is successfully set with " &strArrArg(3)
						If htmlreport = "1" Then
							Call Update_log(MAIN_FOLDER, g_sFileName,"pass")	
						End If
						blnXmlFlag=True 
				Exit for
   						End If
				Next
		End If
Next
		oDoc.SaveFile(sPath)'Save the XML file after updating the value
			If blnXmlFlag=False Then
				Reporter.ReportEvent micFail , "The Element not  found " ,"The element  does not exists in XMLFile"
			End If	
Set oChildEle=Nothing
Set oRootEle=Nothing
Set oDoc=Nothing
End Function
'###############################################################################################################

'###############################################################################################################
'Function name 	: Func_emailCore
'Description    : This function is used to  send email
'Parameters     : 
'Assumptions    : NA
'###############################################################################################################
'The following function is for "sendemail"  Keyword 
'###############################################################################################################
Function Func_emailCore()
Dim sSendTo			 'variable stores the send email address
Dim sSubject				'stores the subject of the mail
Dim sSendToCC				 'stores the CC email address
Dim sSendToBCC 					'stores the BCC email address
Dim sBody 										'stores the body of the body
Dim sAttachment 								'stores the attachment  file
Dim wshNetwork  									'QTP object created for accessing networks
Dim strDomain 												 'stores the domain
Dim strSMTPnetwork
Dim objMessage 													'QTP object created to sending messages
Dim sFrom  																	 'stores the from address
Dim ints 																				'variable used for  looping
Dim arrSendTo 																		'stores the no. of address to be sent
Dim arrCC 																						 'concantenate  CC address
Dim arrBCC																							'concatenate  BCC address
Dim sCC  																									'used to store the 'n' number of cc address
Dim sBCC																										 'used to store the address of 'n' number of BCC address
Dim intc  																													'used for  looping
Dim intb 																														'used for looping
sFrom=arrAction(0)
sSendTo=arrAction(1)
sSendToCC=arrAction(2)
sSendToBCC=arrAction(3)
sSubject=arrKeyValue(0)
sBody=arrKeyValue(1)
sAttachment=arrKeyValue(2)
arrSendTo = Split(arrAction(1),",") 
If ubound(arrSendTo) > "0" Then
	For ints = 0 to ubound(arrSendTo)
		sSendTo = sSendTo&";"&arrSendTo(ints)
   	Next
Else
	sSendTo = arrAction(1)
End If
arrCC = Split(arrAction(2),",") 
If ubound(arrCC) > "0" Then
	For intc = 0 to ubound(arrCC)
		sCC = sCC&";"&arrCC(intc)
   	Next
Else
	sCC = arrAction(2)
End If
arrBCC = Split(arrAction(3),",") 
If ubound(arrBCC) > "0" Then
	For intb = 0 to ubound(arrBCC)
		sBCC = sBCC&";"&arrBCC(intb)
       	Next
Else
	sBCC = arrAction(3)
End If
'
   'Determine the network to which the smtp sever points to
	Set wshNetwork = CreateObject("WScript.Network")
	strDomain = wshNetwork.UserDomain
   If lcase(strDomain) = "in" Then
		strSMTPnetwork = "smtp.keane.com"
	else
		strSMTPnetwork = "" '###################### Required smtp address
   End If
	'initialize all the email stuff
	Set objMessage = CreateObject("CDO.Message") 
	objMessage.Subject = sSubject 
	objMessage.From= sFrom'"xyz123@keane.com"
	objMessage.To = sSendTo 

	If  sSendToBCC<>""Then
	objMessage.BCC = sSendToBCC
	End If

If  sSendToCC<>""Then
	objMessage.CC = sSendToCC
End If
	
	objMessage.TextBody = sBody &  vbCrLf & "DISCLAIMER : This email notification is sent through QTP at "& now &" as part of  Automation Team"
	If not sAttachment = "" then
	objMessage.TextBody = sBody & ": Please refer to the attachment." & vbCrLf & "DISCLAIMER : This email notification is sent through QTP at "& now &"  as part of EnterpriseRx Automation"
	objMessage.Addattachment sAttachment
	End If

	'==This section provides the configuration information for the remote SMTP server.
	'==Normally you will only change the server name or IP.
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
	
	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPnetwork
	
	'Server port (typically 25)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
	
	objMessage.Configuration.Fields.Update
	'Send email 
	objMessage.Send
	If err.number = 0 Then
		Reporter.ReportEvent micPass,"Email Notification to " & sSendTo,sSubject
				If htmlreport = "1" Then
					Call Update_log(MAIN_FOLDER, g_sFileName,"pass")	
				End If
		else
		Reporter.ReportEvent micWarning, "Notification Failure","Email could not be sent to " & sSendTo 
	End If
	Set wshNetwork = nothing
	Set objMessage = nothing
End Function

'##########################################################################################################

'#########################################################################################################
'Function name 	: Func_CallAction
'Description    : This function is used to  used to call a Reusable Action 
'Parameters     : NA
'Assumptions    : NA
'##########################################################################################################
'The following function is for CallAction  Keyword 
'##########################################################################################################
Public Function Func_CallAction()  
	Dim intParcnt		 'stores the ubound of arrKeyParameter
	Dim strActionName 'Stores the Reusable action Name
	Dim tempRowCount 	'stores the row count of datatable
	Dim strActionPath			 'stores reusable action path
	Dim strActionCount 				'stores the ubound of arrAction
	Dim strIteration 						'stores the no. of iteration to be performed, in the action
	Dim Stmnt 										'variabl;e for executing the call action 
	Dim iAct 											'variable used for looping
	Dim qtApp										 'QTP Object  created 
    Dim qtFolders 									'object variable for  accessing folders
	Dim strQTPToolsOptionPath			'variable stores the folder items
	Dim  i  														 'variable used for looping
	Dim ArrKeyPrameter 									'stores the parameter to be passed in the action
	tempRowCount = intRowCount
	strActionCount = Ubound(arrAction)
	strActionName="Action1"'Assumed that the Reusable Action name is "Action1"
	strActionPath=""
    Stmnt =  "LoadAndRunAction strActionPath, strActionName, oneIteration"

Environment("ActionParameterArray")=DataTable.Value(4,dtLocalSheet) 'if parameter is passed in the 4th column of the datatable fo the reusuale action
ArrKeyPrameter=split(Environment("ActionParameterArray"),":")

For i=0 to uBound(ArrKeyPrameter)
ArrKeyPrameter(i)=GetValue(ArrKeyPrameter(i))
Next

strActionPath=GetValue(DataTable.Value(5,dtLocalSheet))'Reusable action path taken from the Fifth column
If strActionPath="" then 'If path is not mentioned in the 5th column then the path is taken fron the QTP>Tools>OptionFolder
	set qtApp = CreateObject("QuickTest.Application") 
			set qtFolders = qtApp.Folders 
			strQTPToolsOptionPath=qtFolders.Item(1)          ' 
			strActionPath=strQTPToolsOptionPath&"\"&"ReusableActions" 'Assumed that all the Reusable action under the "ReusableActions" folder
			set qtApp =Nothing
			strActionPath=strActionPath&"\"&arrAction(0)

Else 
			strActionPath=strActionPath&"\"&arrAction(0)'Path is concanated with the Reusable testname

End If

If strActionCount = "0" then 
		  If Isempty(arrKeyIndex) Then
			   Stmnt =  "LoadAndRunAction strActionPath, strActionName, oneIteration"
		       Execute(Stmnt)	
		else
		       intParcnt = Ubound(ArrKeyPrameter)
			   Stmnt =  "LoadAndRunAction strActionPath, strActionName, oneIteration"
			   For iAct=0 to intParcnt
			    	Stmnt=Stmnt & ", ArrKeyPrameter(" & iAct & ")"
			  Next	
			   Execute Stmnt
	     End If

ElseIf strActionCount = "1" Then

If IsNumeric(arrAction(1)) Then 'checking for the numeric value to identify the no. of iteration
		strIteration=arrAction(1)		
		Stmnt =  "LoadAndRunAction strActionPath, strActionName, oneIteration"

			     If Isempty(arrKeyIndex) Then            
					For i=0 to strIteration-1
			            Execute Stmnt
	                   Next
			    else
			             intParcnt = Ubound(ArrKeyPrameter)
	                     For iAct=0 to intParcnt
				        Stmnt=Stmnt & ", ArrKeyPrameter(" & iAct & ")"
		                 Next
		
	                    For i=0 to strIteration-1
			            Execute Stmnt
	                   Next

		      End If
			Else 
			strActionName = arrAction(1)
		  If Isempty(arrKeyIndex) Then
			   Stmnt =  "LoadAndRunAction strActionPath, strActionName, oneIteration"
		       Execute(Stmnt)	
		else
		       intParcnt = Ubound(ArrKeyPrameter)
			   Stmnt =  "LoadAndRunAction strActionPath, strActionName, oneIteration"
			   For iAct=0 to intParcnt
			    	Stmnt=Stmnt & ", ArrKeyPrameter(" & iAct & ")"
			  Next	
			   Execute Stmnt
	     End If
End If

ElseIf strActionCount = "2" Then
		          strActionName = arrAction(1) 'stores the action name
		          strIteration = arrAction(2)	'stores the no.of iteration 
		         Stmnt="LoadAndRunAction strActionPath, strActionName, oneIteration"
		 If Isempty(arrKeyIndex) Then
				   For i=0 to strIteration-1
				   Execute Stmnt
				   Next
		else
	             intParcnt = Ubound(ArrKeyPrameter)			       		  
	              For iAct=0 to intParcnt
		           Stmnt=Stmnt & ", ArrKeyPrameter(" & iAct & ")"
		          Next  
	           For i=0 to strIteration-1
			   Execute Stmnt
	            Next
		
	End If
Else
		Reporter.ReportEvent micFail,"The call action syntax check","The call action syntax is invalid"
End If
	intRowCount = tempRowCount
End Function
'#########################################################################################################

'#########################################################################################################
'Function name 	: Func_CallNestedAction
'Description    : This function is used to  used to call a Reusable Action from the Reusable action
'Parameters     : The ReusableTestName, Action name,Iteration and the parameters for the reusable action.
'Assumptions    : NA
'##########################################################################################################
'The following function is for CallNestedAction  Keyword 
'##########################################################################################################
Public Function Func_CallNestedAction()

	Dim intNestedParcnt
	Dim strNestedActionName 'Stores the Reusable action Name
	Dim tempNestedRowCount
	Dim strNestedActionPath
	Dim strNestedActionCount
	Dim strNestedIteration
	Dim strNestedStmnt
	Dim iNestedAct
	Dim qtNestedApp
    Dim qtNestedFolders
	Dim strNestedQTPToolsOptionPath
	Dim intarr
	Dim ArrNestedActionKeyPrameter

	tempNestedRowCount = intRowCount
	strNestedActionCount = Ubound(arrAction)
	strNestedActionName="Action1"'Assumed that the Reusable Action name is "Action1"
	strNestedActionPath=""
    strNestedStmnt =  "LoadAndRunAction strNestedActionPath, strNestedActionName, oneIteration"
	If  DataTable.Value(4,dtLocalSheet)<>""Then
			Environment("ActionNestedParameterArray")=DataTable.Value(4,dtLocalSheet)
		ArrNestedActionKeyPrameter=Split(Environment("ActionNestedParameterArray"),":")
		If isarray(ArrNestedActionKeyPrameter) Then
	For intarr=0 to uBound(ArrNestedActionKeyPrameter)
	ArrNestedActionKeyPrameter(intarr)=GetValue(Lcase((ArrNestedActionKeyPrameter(intarr))))

	Next
	  intNestedParcnt = Ubound(ArrNestedActionKeyPrameter)
		End If
		else 
		intNestedParcnt=0
  End If

strNestedActionPath=GetValue(DataTable.Value(5,dtLocalSheet))'Reusable action path taken from the Fifth column
If strNestedActionPath="" then 'If path is not mentioned in the 5th column then the path is taken fron the QTP>Tools>OptionFolder
	set qtNestedApp = CreateObject("QuickTest.Application") 
			set qtNestedFolders = qtNestedApp.Folders 
			strNestedQTPToolsOptionPath=qtNestedFolders.Item(1)          '
			strNestedActionPath=strNestedQTPToolsOptionPath&"\"&"ReusableActions" 'Assumed that all the Reusable action under the "ReusableActions" folder
			set qtNestedApp =Nothing
			strNestedActionPath=strNestedActionPath&"\"&arrAction(0)

Else 
			strNestedActionPath=strNestedActionPath&"\"&arrAction(0)'Path is concanated with the Reusable testname

End If

If strNestedActionCount = "0" then 
		  If intNestedParcnt=0 Then
		       strNestedStmnt =  "LoadAndRunAction strNestedActionPath, strNestedActionName, oneIteration"
		       Execute(strNestedStmnt)	
		else
		   
			   strNestedStmnt =  "LoadAndRunAction strNestedActionPath, strNestedActionName, oneIteration"
			   For iNestedAct=0 to intNestedParcnt
			    	strNestedStmnt=strNestedStmnt & ", ArrNestedActionKeyPrameter(" & iNestedAct & ")"
			  Next	
			   Execute strNestedStmnt
	     End If

ElseIf strNestedActionCount = "1" Then
		strNestedIteration=arrAction(1) 'stores the no. of iteration
		
		strNestedStmnt =  "LoadAndRunAction strNestedActionPath, strNestedActionName, oneIteration"

			     If intNestedParcnt=0 Then            
			             Execute strNestedStmnt
			    else
			      
	                     For iNestedAct=0 to intNestedParcnt
				        strNestedStmnt=strNestedStmnt & ", ArrNestedActionKeyPrameter(" & iNestedAct & ")"
		                 Next
		
	                    For intarr=0 to strNestedIteration-1
			            Execute strNestedStmnt
	                   Next

		      End If

ElseIf strNestedActionCount = "2" Then
		          strNestedActionName = arrAction(1)'stores the action name
		          strNestedIteration = arrAction(2)	'stores the no. of iteration
		         strNestedStmnt="LoadAndRunAction strNestedActionPath, strNestedActionName, oneIteration"
		 If intNestedParcnt=0 Then
				   for intarr=0 to strNestedIteration-1
				   Execute strNestedStmnt
				   Next
		else			       	  
	              For iNestedAct=0 to intNestedParcnt
		           strNestedStmnt=strNestedStmnt & ", ArrNestedActionKeyPrameter(" & iNestedAct & ")"
		          Next  

	           for intarr=0 to strNestedIteration-1
			   Execute strNestedStmnt
	            Next
		
			End If
Else
		Reporter.ReportEvent micFail,"The call nested action syntax check","The call nested action syntax is invalid"
End If
	intRowCount = tempNestedRowCount
End Function
'##########################################################################################################

'##########################################################################################################
'Function name 	: Func_CaptureData
'Description    : It is used to pass test data between scripts at runtime
'Parameters     :NA
'Assumptions    : Specified file is is an Excel work book
'##########################################################################################################
'The following function is used for keyword 'CaptureData'.
'##########################################################################################################
Function Func_CaptureData()
Dim strSheetName 'stores the name of thr sheet 
Dim objExcel    			 'Holds the excel object 
Dim objWorkbook  		'Holds the work book object
Dim countnt      					'stores the number of sheets 
Dim ArrOTestName			 'stores the details of the test data sheet
Dim arrIOParam  			    	 'Stores the values of Input & Output parameters 
Dim OParamName  				 'Stores the value of Out put parameter Name
Dim OParamVal                           'Stores the value of Out put parameter Value
Dim iOCount                                    'Counter for looping
Dim strVarName                              'Stores the name of the variable in Data sheet 
Dim i		                                             	'used for looping
Dim OTestPath                                     'path of the excel sheet where it is located
Dim flag                                                    'stores flag value
Dim sheetname                                      'stores the excel sheet name
Dim objSheet                                            'object variable used to access the excel sheet
Dim newcol		                                            'variable stores the used column count of the excel sheet
Dim newrow		                                             'variable stores the used row count of the excel sheet
Dim OTestName                                            'sheetname of the excel sheet 
	OParamName = arrAction(0) 'stores the header name of the column name
	OParamVal = arrAction(1)  'stores the value to be updated in the specific column
	If Ubound(arrkeyvalue)>0 Then
			If Instr(arrkeyvalue(1),"xls")>0 Then
				OTestPath=arrKeyValue(0) &":"& arrKeyValue(1)
			Else
				OTestPath=arrKeyValue(0)
				OTestName=arrKeyValue(1)
			 End if 
	Else
		OTestPath=arrKeyValue(0)
	End If

If Ubound (arrKeyValue)=2  Then
OTestName=arrKeyValue(2)
End If

    Set objExcel = createObject("EXCEL.Application")
 If OTestPath <> "" Then
    Set objWorkbook = objExcel.Workbooks.Open(OTestPath)
	countnt = objExcel.ActiveWorkbook.sheets.count
    	flag =0
	'checking for the sheet with testcase name to be present
	If  OTestName <> ""Then

	For i = 1 to countnt
				sheetname = objExcel.ActiveWorkbook.sheets(i).name
				If LCase(Trim(sheetname)) = LCase(Trim(OTestName)) Then
							flag = 1
							Exit for
				End If
	Next
	Else
		sheetname=objExcel.ActiveWorkbook.sheets(1).name 'getting the name of the excel sheet
		flag=1
	End If
	
	If flag = 0 Then
				Reporter.reportevent micWarning, "The specified sheet for the test script does not exist.”,” So, added the sheet"
				.objExcel.Application.Quit
				Set objExcel=Nothing
				Exittest
	End If
	
	'activating the sheet for the testcase
	Set objSheet = objExcel.ActiveWorkbook.Worksheets(sheetname)
	objExcel.visible=false
	newcol = objSheet.usedrange.columns.count
	newrow = objSheet.Usedrange.rows.count
	'writing the variable name and value to the desired cell
	For iOCount = 1 to newcol
		strVarName = objSheet.Cells(1,iOCount).value 'getting the columns titles
		If LCase(Trim(strVarName)) = LCase(Trim(OParamName)) Then
			Exit for
		End If
	Next
	If iOCount > newcol Then
		objSheet.Cells((1),iOCount).value = OParamName
	End If
    objSheet.Cells((intDataCounter+1),iOCount).value = OParamVal 'given value from datatable is set into the specified sheet column of the excel sheet
	objExcel.ActiveWorkBook.Save
	objExcel.ActiveWorkBook.Close
	Reporter.ReportEvent micPass, "To add the given value from datatable into the specified  column of the excel sheet", "The value ' " & OParamVal &"' is set to  the column name  -  " & OParamName
 Else 
	Reporter.ReportEvent micWarning,"Test Data sheet Path not specified", "Unable to pass Parameter to the Specified Script"	
 End If
	objExcel.Application.Quit
	Set objExcel=Nothing
End Function
'##########################################################################################################

'##########################################################################################################
'Function name	: DebugFunc
'Description	: If User requires to debug the keyword script then the user can use this function
'Parameters	: NA
'Assumptions 	: All the Environment variables are not stored in a XML file and then loaded when debuggin because
'		  the QTP script might have an existing XML file associated.
'##########################################################################################################
'The following function is for Debug Keyword
'##########################################################################################################
Function Func_Debug()
	Dim intStartRow 'takes the start row number from datatable , where which the debug should start
	Dim intEndRow   ''takes the end row number from datatable , where which the debug should end
	'Define the values for Global Variables
	Environment("intStartRow")=arrAction(0)
	 Environment("intEndRow")=arrAction(1)
	Environment("LogFile")=arrKeyValue(0)
	Environment("PrintOption")=arrKeyValue(1)
End Function
'##########################################################################################################

'##########################################################################################################
'Function name	: DebugGetEnv
'Description	: If User requires to debug the keyword script then the user can use this function
'Parameters	:NA
'Note :It update the file or print log for storekeyword and check
'##########################################################################################################
'The following function is is called from the Func_Store_Java function
'##########################################################################################################
Public Function DebugGetEnv()
	Dim strText				'As the contents of the log file
	Dim ostrFile			'As the file object
	Dim oFSO				'As the File system object
	Dim strDesktop		 'As the Path of the desktop
	Dim oQtApp			    'As the QuickTest Object
	Dim oWshShell		  'As the Windows Shell object
	Dim strContents		    'As the contents of the log file
	Dim arrSplit2			     'As the Array
	Dim strFileName		    'As the file name
	Dim EnvSplit			      'As the interim array
	Dim VariableName	'As the variable name
	Dim oFile				          'object used to creat a text file
	Dim checkvalsplit           'variable stores the 5th column of the datatable
	Dim checkval 		            'variable is  initialized as "0" or "1" , according the the pass or fail condition
	Dim sfileopen			          'initialization variable
	Dim strIntPath
   	strIntPath = "C:\O2T\"  
	'If the log file is required for the debugging, then load all the Environment variables to the QTP
	If Environment("LogFile")="true" Then
		'Create a Shell object to access the desktop
		Set oWshShell = CreateObject("WScript.Shell")
	'Store the value of the desktop Path
		strDesktop = oWshShell.SpecialFolders("Desktop")
		'Create a File system object
       Set oFSO = CreateObject("Scripting.FileSystemObject")
	   	'Create a Quick test application object to access the test name
		Set oQtApp = CreateObject("QuickTest.Application")
	
		If Not oFSO.FolderExists(strIntPath) Then
			Set objMain = oFSO.CreateFolder(strIntPath)
		End If
		'Store the file name in a variable
            strFileName = "C:\O2T\"& oQtApp.Test.Name &"Log.txt"
	'If the File exist then load the variables in QTP
	If not oFSO.FileExists(strFileName) Then
        Set oFile =oFSO.CreateTextFile(strFileName,True)
		sfileopen=1
	End If
		If initial="storevalue" Then
			VariableName = arrKeyValue(1)
		  Else
		  checkvalsplit = Datatable.Value(5,dtLocalSheet)
			If Instr(checkvalsplit,":") >0 Then
					checkval = Split(checkvalsplit, ":", -1, 1)
			  		  VariableName = lcase(checkval (0))
			Else
				  VariableName = checkvalsplit	
		  End If

		End If
		
		If Environment(lcase(VariableName))<>"" Then
			strText = (strText) &VBCR&"Environment Variable Name ="& VariableName &vbTab&"Value ="&Environment.Value(lcase(VariableName))
		End If
		If Environment("PrintOption")="true" Then 'if stated as true in datatable, print option will be done
			Print strText
		End If
	  If sfileopen <>1  Then
		Set ostrFile= oFSO.OpenTextFile(strFileName, 8)
		ostrFile.WriteLine strText
	   	ostrFile.Close
	  Else
		oFile.WriteLine strText
		ofile.Close
	  End If
End If
			'Clear all object variables
      Set oQtApp = Nothing
	  Set oWshShell =Nothing
	  Set oFSO = Nothing
	  Set ostrFile= Nothing
End Function
'###########################################################################################################

'###########################################################################################################
'Function name	: Func_ScreencaptureOption()
'Description	: If User requires to take the screenshot  for Check,Perform and Context.
'Parameters	: None
'###########################################################################################################
Function Func_ScreencaptureOption()
			Dim iperform 'variable intialized and incremented to 1,  for the screen capture  for the operation "perform" 
			Dim icontext 'variable intialized and incremented to 1,  for the screen capture  for the operation "context" 
			Dim icheck   'variable intialized and incremented to 1,  for the screen capture  for the operation "check" 
			Dim opt 	 'Used for looping
			Dim intp 	 'Getting the values from 3rd column of datatable and concatening with "i"
			If objname="" Then
			Environment.Value("iperform")=1
			Environment.Value("icontext")=1
			Environment.Value("icheck")=1
			Else 
			For opt=0 to Ubound(arraction)
			intp="i"&arraction(opt)
				If LCase(intp)="iperform" Then
					Environment.Value("iperform")=1
					Elseif LCase(intp) ="icontext" Then
					Environment.Value("icontext")=1
					Elseif LCase(intp) ="icheck" Then
					Environment.Value("icheck")=1
					Else
						Reporter.ReportEvent micFail, "option does not exist", "Option "&intp &" does not exist"
				End If
			Next
			End If
	End Function
'###############################################################################################################

'###############################################################################################################
'Function name  : func_tablesearch
'Description    : This function is used to search for the particular row and Column in the table based on the  multiple search '              
' criteria entered in keyword script for perform keyword.
'                This function is used to Check for text present in a table  based on the  multiple
'                criteria entered in keyword script by Check and  to return the column number in perform keyword .
'Parameters     : object is a current object on which action should to be performed or checked
'                strsearch is a multiple search criteria entered in keyword script
'                (ex.<colname1>,<colValue2>::<colname1>,<colValue2>...........)
' If the value to be returned then use with Check keyword by passing the columns name  in 5th column
'                (ex:stroutcol1:stroutcol2:stroutcol3 .... are  used to store the output  columns values and return the row number as well )
'If the Row number to be returned then use with Perform keyword passing the columns name in fourth column and variable name in 5th column.                

'Assumptions    : NA
'################################################################################################################
Public Function Func_tablesearch(Object,strsearch)

Dim arrindex 						'Stores the arr index
Dim blnSearchcolfound		'Flag to store the column is found
Dim intoutcolvalues					 'Store the out column index
Dim strcolname 								'Store the column name
Dim intCols											'Store total number of columns
Dim intindexcol										'Stores the column index
Dim strExpectedColValues				'Stores the Expected search after concatinating
Dim introws													'Stores the total number of rows
Dim introw														'Index to each row
Dim blnRowFound											'Holds the flag if the row is found
Dim strActColValues										'Stores the Actual search string after concating all
Dim intarrrowindex												'Point the arra index
Dim introwfound														'Store the row  number found
Dim strActCellVal													'Stores the cell value
Dim arrColIndex 														'stores the column number
Dim arrColsValue 														'column value to be searched in table
Dim arrCol 																			'array variable stores the data table value from 4th column which splits with delimiter "--"
Dim Colcount 																		' Takes the column count specified in datasheet explicitily
Dim strte																					 'Takes the column count of the table

strsearch=GetValue(strsearch)
If instr(strsearch,"--") >0 then 'Loop is used to search ,between the given column index which is customized with values.
	arrCol = Split(strSearch,"--")
	Colcount=arrCol(1)
	arrsearchcriteria=split(arrcol(0),"::") 'spliting all individual  column names with values
	 strte = Colcount-1
Else
arrsearchcriteria=split(strsearch,"::")
End If
ReDim arrColIndex(ubound(arrsearchcriteria))
ReDim arrColsValue(ubound(arrsearchcriteria))
arrindex=0
blnSearchcolfound=False

For each arreachsearch in arrsearchcriteria
arreachsearch=split(arreachsearch,",")
arreachsearch(0)=GetValue(Cstr(arreachsearch(0)))
arreachsearch(1)=GetValue(Cstr(arreachsearch(1)))
If Colcount = "" then
	intCols= Object.GetROProperty("cols") 'getting the no. of columns in the table
	strte = intCols-1
End If
For intindexcol=0 to strte
			strcolname=Object.Getcelldata(1,intindexcol) 'storing the column name
               If Trim(lcase(strcolname))=Trim(lcase(arreachsearch(0))) Then'checking whether the column name given in datatable equals the coloumn namtable in application
                blnSearchcolfound=True
                Exit for
              End If
			  Next
  If blnSearchcolfound=False Then
  Reporter.ReportEvent micFail,"Column "&arreachsearch(0)&" is not found"," Please recheck the Column description. Cannot continue further execution"
  	intRowCount=DataTable.LocalSheet.GetRowCount
  ExitTest()
  End If

 If blnSearchcolfound=True Then
 arrColIndex(arrindex)=intindexcol 
 arrColsValue(arrindex)=arreachsearch(1)
 arrindex=arrindex+1
 End If
Next

strExpectedColValues=Join(arrColsValue)
strExpectedColValues=Trim(Replace(strExpectedColValues," ","",1))
introws=Object.GetROProperty("rows")
blnRowFound=False
For introw=0 to introws-1
strActColValues=""
              For intarrrowindex=0 to Ubound(arrColIndex)
                 strActCellVal=Object.GetCellData(introw,arrColIndex(intarrrowindex))
                 strActColValues=Trim(Replace(strActColValues&strActCellVal," ","",1))
             Next
If Lcase(strActColValues)=Lcase(strExpectedColValues) Then
	introwfound=introw
	blnRowFound=True
	Func_tablesearch=introw&":"&strActColValues&":"&strExpectedColValues
Exit for
End If
Next

If blnRowFound="False" Then
introw=" "
Func_tablesearch=introw&":"&strActColValues&":"&strExpectedColValues
End If
'#############Keyword is Perform##########
If initial="perform" Then
Dim  strOutCols 'stores the value from 5th column of datatable
Dim  intindexColIndex  'variable used for arithmetic calculation
Dim strOutColsNames  '5th column value get through the function Func_GetValue
    If  DataTable.Value(5,dtLocalSheet)<>"" Then
    strOutCols=CStr(Trim(DataTable.Value(5,dtLocalSheet)))
    strOutColsNames=GetValue(strOutCols)
    arrOutColsvars=Split(strOutCols,":")'Store the variable name to store the columns number in an array
   End If
 varName=Lcase(Trim(arrOutColsvars(0)))
 Environment(varName)=introwfound
For intoutcolvalues=1 to ubound(arrOutColsvars)
 intindexColIndex= intoutcolvalues-1
varName=Lcase(Trim(arrOutColsvars(intoutcolvalues)))
Environment(varName)=arrColIndex(intindexColIndex)
Next
Reporter.ReportEvent micDone, "To get the Row numbers of the given search criteria","The given search criteria '" &strExpectedColValues&"'  exists in row number  '" &introwfound &"'"
End If
End Function
'#####################################################################################################

'#####################################################################################################
'Function name 		: Func_SelectText
'Description        : If the user wants to click a text in any windows, dialog, this function is used
'Parameters       	: Text to be clicked should be send as parameter
'Return Value		:NA
'Assumptions     	: NA
'#####################################################################################################
'The following function is used for textclick Keyword
'#####################################################################################################
Function Func_SelectText(text)
	Dim strWholeWord 'Holds the boolean value
	Dim intl'variable for initialization
	Dim intt 'variable for initialization
	Dim intr  'stores the width property of an object at run time
	Dim intb 'stores the height property of an object at run time
	Dim Success 'stores the boolean value as it checks with the existance of the specified keyword in the object
	Dim strRetry 'variable used for looping
	Dim hl  'temporary storing variables
	Dim ht   'temporary storing variables
	Dim hr 'stores the width property of an object at run time
	Dim hb  'stores the height property of an object at run time
	strWholeWord = "False"
	If (Ubound(arrKeyIndex) > 1) Then
		strWholeWord = arrKeyIndex(2) 'This value can be True/False
	End If
	intl = 2
	intt = 2
	intr = CInt(object.GetROProperty("width")) - 10
	intb = CInt(object.GetROProperty("height")) - 10
	hl = intl
	ht = intt
	hr = intr
	hb = intb
	strText = object.GetVisibleText(intl, intt, intr, intb)
    	Success = object.GetTextLocation(text, intl, intt, intr, intb, strWholeWord)
	'If the initial search for text is not successful, put in a means to expand the search by a pixel in each the top and bottom coordinate.
	For strRetry = 1 to 3
		If Success Then
			Reporter.ReportEvent micDone, "TextClick", "Text '" & text & "' Found."& vbCrLf &" Displayed text:  "  & vbCrLf & strText & vbCrLf & "Retries:  " & (retry - 1)
			object.Click CInt((intl+(intr-ht+hl))/2), CInt(((intt-ht+hl)+intb)/2)
			Exit For
		Else
			intl = hl
			intt = ht - strRetry
			intr = hr
			intb = hb + strRetry
			strText = object.GetVisibleText(hl, ht, hr, hb)
			Success = object.GetTextLocation(text, intl, intt, intr, intb, strWholeWord)
		End If
	Next
	If Not Success Then
		Reporter.ReportEvent micFail, "TextClick", "Text Not Found - " & strText & vbCrLf & "Text Displayed:  " & vbCrLf & text & vbCrLf & "Retries:  " & (retry - 1)
	End If
End Function
'##################################################################################################################