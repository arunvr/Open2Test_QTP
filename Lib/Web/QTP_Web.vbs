'###############################  KEYWORD DRIVEN FRAMEWORK   #####################################
'Project Name		: Web Framework
'Author		       	: Open2Test
'Version	    	: V 2.1
'Date of Creation	: 12-Mar-2013
'######################################  Driver Function  ##################################################
'#################################################################################################
'Function name 	: Keyword_Web
'Description             : This is the main function which interprets the specific keywords and performs the desired actions. All the specific keywords 
'                                      used in the datatable are processed in this function.
'Parameters           : Keyword in the 2nd column of the data table
'Assumptions     	: The Automation Script is present in the Local Sheet of QTP.
'#################################################################################################
'The following function is for 'Context' keyword.
'#################################################################################################
Function Keyword_Web(initial)
   	 If htmlreport = "1" Then
		Call Update_Log(MAIN_FOLDER, g_sFileName,"executed")' calling function update log to create an execution log in HTML file
	 End If	
   On error resume next
   If initial = "context" Then
	   Call Func_Context_Web(arrAction,intRowCount)
	   Exit Function
   End If
    Set object = Nothing
	Call Func_ObjectSet_Web(arrAction,intRowCount)
            'Perform and check operations
	Select Case LCase(initial) 'to perform keyword operation given from datatable
		Case "perform"
			'start perform
			Call Func_Perform_Web(object,arrAction,arrKeyValue,arrKeyIndex,intRowCount)
        Case "storevalue"
			Call Func_Store_Web(object,arrAction)
            'Checking
		Case "check"
			Call Func_Check_Web(object,arrAction,arrKeyValue,arrKeyIndex,intRowCount)
		Case "dragdrop"
			Call Func_DragDrop_Web(arrKeyValue,arrKeyIndex,object) 'call to drag drop function 
        Case Else
			errStr1 = "Keyword check at Line No. -" &intRowCount
			errStr2 = "The keyword -'" & initial & "' is not supported.Please verify the keyword entered."
			err.raise vbObjectError	
		End Select 'End of perform,storevalue, Checking	
End Function

'#################################################################################################
'Function name 	: 	Func_Context_Web
'Description   		 : This function is used to set the full hierarchical Path for the object on which 
'									some action is to be performed.
'Parameters     	: The Object details as the full hierarchical Path of the Object goes as parameter 
'									to the function and the current row number in the Local Sheet.           
'Assumptions   	 : AUT is already up and running. 
'#################################################################################################
'The following function is for 'Context' keyword.
'#################################################################################################
Function Func_Context_Web(arrObj,intRowCount)
	Dim strReportData  	'Stores the contents of the fourth Column of current row in Global Sheet
	Dim arrChildCell	'stores the elements separated by the delimiter '::'
	Dim arrFramed		'Stores the elements according to the object type and name
	Dim contextData		'Stores the value present in the fourth Column
	Dim arrChild		'Stores the child objects of the main window
		If (CStr(Trim(DataTable.Value(4,dtLocalSheet))) <> "") Then
			strReportData = CStr(Trim(DataTable.Value(4,dtLocalSheet)))
			arrChildCell = Split(strReportData, "::", -1, 1)
			FrameName =""
			inti =0
			For inti = 0 to Ubound(arrChildCell)
				arrFramed = Split(arrChildCell(inti),";",2,1)
				FrameName = arrFramed(1) 'Storing the current screen name
			Next
       End If 
	Call Func_DescriptiveObjectSet(arrObj,intRowCount)
	Select Case LCase(arrObj(0)) 'sets the hierarchical path of the object
		Case "browser" 'Parent is Browser
			Set curParent = Browser(arrObj(1)) 'initial declaration of the parent object
		Case "swfwindow" 'parent is swfwindow
			Set curParent = SwfWindow(arrObj(1)) 'used to check the current browser
		Case "window" 'Parent is Window (e.g. error messages)
			Set curParent = Window(arrObj(1)) 'initial declaration of the parent object
		Case "dialog" 'Parent is dialog
			Set curParent = Dialog(arrObj(1)) 'initial declaration of the parent object
	End Select
	If (CStr(Trim(DataTable.Value(4,dtLocalSheet))) <> "") Then
		contextData = CStr(Trim(DataTable.Value(4,dtLocalSheet)))
		arrChildCell = Split(contextData, "::")
		For intj = 0 To UBound(arrChildCell)
			arrChild = Split(arrChildCell(intj), ";", 2, 1) 'checks the child object type
			arrChild(1) = GetValue(Trim(arrChild(1)))
			Dim arrChildDesc
			arrChildDesc = Split(arrChild(1),":=",-1,1)
			If UBound(arrChildDesc) > 0 Then
				Call Func_DescriptiveObjectSet(arrChild,intRowCount)
			End If
			Select Case LCase(arrChild(0)) 'Child object types
				Case "dialog"
					Set curParent = curParent.Dialog(arrChild(1))
				Case "page"
					Set curParent = curParent.Page(arrChild(1))
				Case "frame"
					Set curParent = curParent.Frame(arrChild(1))
				Case "table"
					Set curParent = curParent.WebTable(arrChild(1))
				Case "window"
					Set curParent = curParent.Window(arrChild(1))
				Case "swfwindow"
					Set curParent = curParent.SwfWindow(arrChild(1))
				Case Else
					 errStr1 = "Keyword Check at Line no - " & intRowCount
					 errStr2 = "Keyword - '" & arrChild(0) & "'  not supported. Please verify Keyword entered."
					 err.raise vbObjectError
			End Select
		Next
		newContext=1
		parChild = arrChild(0)'setting to check whether the parent is web or win
	End If
	Set parent = curParent 'Setting the current screen under which the object is present
			If  Environment.Value("icontext")=1 Then
		Call Func_CaptureScreenshot("test",intRowCount) 'call screencapture function to take screenshot
		End If
End Function
'#################################################################################################
'#################################################################################################
'Function name 	  : Func_ObjectSet_Web
'Description      : This function sets the parent and child objects.
'Parameters       : arrObj is an array of object names and intRowCount is the current Row number.
'Assumptions      : NA
'#################################################################################################
'The following function is called Internally 
'#################################################################################################
Function Func_ObjectSet_Web(arrObj,intRowCount)
   On error resume next
   Dim ObjectVal 'Stores the Table object Class name
   Call Func_DescriptiveObjectSet(arrObj,intRowCount)   
   'If condition for object setting for objects other than 'Window', 'Dialog' and 'Browser'
	If lcase(arrObj(0))<>"window" And arrObj(0) <> "split" And arrObj(0) <> "random" And arrObj(0)<>"dialog" And LCase(arrObj(0))<>"browser" And arrObj(0)<>"sqlvaluecapture" And arrObj(0)<>"sqlexecute" And arrObj(0)<>"sqlcheckpoint" Then 
		Select Case LCase(arrObj(0)) 'setting of  objects
			Case "frame"
				Set object = parent.frame(arrObj(1))  'initial declaration of the object
			Case "page"
				Set object = parent.Page(arrObj(1))     
			Case "textbox"
				If Lcase(parChild) = "page" Then
					Set object = parent.WebEdit(arrObj(1))
				Else
					Set object = parent.WebEdit(arrObj(1))  
				End If
			Case "button"
				Select Case LCase(parChild) 'to handle the windows error messages
					Case "frame"
						Set object = parent.Webbutton(arrObj(1))
					Case "page"
						Set object = parent.WebButton(arrObj(1))
					Case "dialog"
						Set object = parent.WinButton(arrObj(1))
				End Select
			Case "wcombobox"
				Set object = parent.Wincombobox(arrObj(1))
			Case "combobox"
				Set object = parent.WebList(arrObj(1))
			Case "checkbox"
				Set object = parent.WebCheckbox(arrObj(1))
			Case "radiobutton"
				Set object = parent.WebRadioGroup(arrObj(1))
			Case "image"
				Set object = parent.Image(arrObj(1))
			Case "table"
				Set object = parent.WebTable(arrObj(1))
			Case "element"
				Set object = parent.WebElement(arrObj(1))
			Case "link"
				Set object = parent.Link(arrObj(1))
			Case "viewlink"
				Set object = parent.Viewlink(arrObj(1))
			Case "webarea"
				Set object = parent.WebArea(arrObj(1))
			Case "webfile"
				Set object = parent.WebFile(arrObj(1))
			Case "menu"
				Set object = parent.WinMenu(arrObj(1))
			Case "toolbar"
				Set object = parent.WinToolbar(arrObj(1))
			Case "tabletextbox"
				ObjectVal = "WebEdit"
			Case "tablebutton"
				ObjectVal = "WebButton"
			Case "tablecombobox"
				ObjectVal ="WebList"
			Case "tablecheckbox"
				ObjectVal = "WebCheckBox"
			Case "tableradiobutton"
				ObjectVal = "WebRadioGroup"
			Case "tableimage"
				ObjectVal = "Image"
			Case "tablelink"
				ObjectVal = "Link"
			Case "tableelement"
				ObjectVal = "WebElement"
			Case "childtable"
				ObjectVal = "WebTable"
			Case Else
				errStr1 = "Keyword Check at Line no - " & intRowCount
				errStr2 = "Keyword - '" & arrObj(0) & "'  not supported.Please verify Keyword entered."
				err.raise vbObjectError
		End Select 'End of object settings
		If Instr(1,Lcase(arrObj(0)),"table") <> 0 And len(arrObj(0)) >5 Then
			If arrKeyValue (0) = "click" Or arrKeyValue(0) = "submit" Then
				childCount = parent.WebTable(arrObj(1)).childItemCount(CInt(arrKeyIndex(1)), CInt(arrKeyIndex(2)),ObjectVal)
			Else
				childCount = parent.WebTable(arrObj(1)).childItemCount(CInt(arrKeyIndex(2)), CInt(arrKeyIndex(3)),ObjectVal)
			End If
			If childCount <> 0 Then
				If arrKeyValue (0) = "click" Or arrKeyValue(0) = "submit" Then
					If ubound(arrKeyIndex) > 2 Then
						index = CInt(arrKeyIndex(3))
					Else
						index = 0
					End If
					Set object = parent.WebTable(arrObj(1)).childItem(CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2)),ObjectVal,index)
				Else
					If ubound(arrKeyIndex) > 3 Then
						index = CInt(arrKeyIndex(4))
					Else
						index = 0
					End If
					Set object = parent.WebTable(arrObj(1)).childItem(CInt(arrKeyIndex(2)),CInt(arrKeyIndex(3)),ObjectVal,index)
				End If
			Else
				Reporter.ReportEvent micFail, "Child Object of  " & arrObj(0) & "-" & "'" &arrObj(1) & "'" & " should be present ","Child Object of  " & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " was not found"
			End If
		End If	
	  Call Func_Wait_Web(arrObj,arrKeyValue,initial)	
	End If 'End of the setting object if condition
End Function
'##############################################################################################
'##############################################################################################
'Function name 		: Func_Perform_Web
'Description        : If the user requires to perform a set of operations then the user can use this function
'Parameters       	: 	1. Object on which the specified operation needs to be performed
'				  					  2. The operation that needs to be performed on the object                                 
'				  	 				 3. Additional parameters if required to identify the object where operation needs to be performed.                                
'Assumptions     	: NA
'##############################################################################################
'The following function is for Perform Keyword
'##############################################################################################
Function Func_Perform_Web(object,arrObj,arrKeyValue,arrKeyIndex,intRowCount)
   On error resume next
   Select Case LCase(arrKeyValue(0))' Selecting the specific action to be performed
		Case "rownum"
		   	    strParam = CStr(Trim(DataTable.Value(5,dtLocalSheet)))
			Call func_getRowNum_Web(object,arrKeyValue(1),strParam)
		Case "tablesearch"
				Call Func_tablesearch(object,arrTableindex(1))
	     Case "text"
			Dim Text 'Stores the text from the specified row and Column
			Dim Row 'Stores the number of rows in the table
			Dim col 'Stores the number of columns in the table
			Dim strstore 'This variable is used to store the data in 4th Column of DataSheet
			Dim strvari 'This variable is used to store the split array  of 4th Column of DataSheet
			Dim checked ' This variable is used to store the ROProperty "value" of specified object
			If arrObj(0) = "table" Or arrObj(0) = "childtable" Then 'Checking if object is table or childtable
				propertyVal = "null"
				strCellData = arrKeyIndex(1)
				Call GetValue(strCellData)
				Row = parent.WebTable(arrObj(1)).RowCount
				For intj = 1 To Row
					col = parent.WebTable(arrObj(1)).ColumnCount(j)
					For inti = 1 To col                                                 
						Text = Trim(parent.WebTable(arrObj(1)).getCellData(j, i))
						If Text = Trim(strCellData) Then 'Comparing the actual and the expected values
							keyword = 0
						   Exit For
						Else
							keyword = 1
						End If
					Next
					If keyword = 0 Then
						Exit For
					End If
				Next
				If keyword = 0 Then
					   strstore = Cstr(Trim(DataTable.Value(4,dtLocalSheet)))
					strvari = Split(strstore,";",2,1)
					Environment.Value(lcase(strvari(0))) = j
					Environment.Value(lcase(strvari(1))) = i
				Else
					Reporter.reportevent micFail, strCellData  & " field should be present", strCellData & " field is not present"
				End If
			ElseIf arrKeyValue(1)="blank" Then
				propertyVal = "null"
				checked=object.GetROProperty("value")
				If checked="" Or checked ="#0" Then
					Reporter.ReportEvent micPass, arrObj(1)&" field should be blank",arrObj(1)&" field is blank"
				Else
					Reporter.ReportEvent micFail, "Keyword Check at Line no.- " & intRowCount &" . '" & arrObj(1)&"'  field should be blank", arrObj(1)&" field is not blank"
				End If
			ElseIf arrObj(0) = "element" Or arrObj(0) = "tableelement" Then
				propertyVal = "innertext"
			Else
				propertyVal = "value"
			End If
		Case "setsecure"
			arrKeyIndex(1) = GetValue(arrKeyIndex(1))
			object.SetSecure (arrKeyIndex(1))
			Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		Case "set"
			Dim propSplit1 'This variable is used to store the split array  of 4th Column of DataSheet when delimiter '_' is used
			Dim strstatus  'This variable is used to store the split array  of propSplit1(0) when delimiter '_' is used
			If (Instr(1,arrKeyIndex(1),"d_") <> 0) Then
				propSplit1 = Split(arrKeyIndex(1),"_",-1,1)
				strstatus = Split(propSplit1(1),";",-1,1)
				VarName = strstatus(0)
				Select Case LCase(strstatus(0)) 'Selecting the appropriate date and time format
					Case "currenttime"
						Environment.Value(lcase(VarName)) = FormatDateTime(Now(),4)
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" - "&arrStatus(0) &" performed successfully on "&arrObj(0)
					Case "currentdate"
						Environment.Value(lcase(VarName)) = FormatDateTime(Now(),2)
						propSplit = Split(Environment.Value(lcase(VarName)),"/",-1,1)
						If  propSplit(0) < 10 Then
							propSplit(0) = 0 & propSplit(0)
							Environment.Value(lcase(VarName)) = propSplit(0) & "/" & propSplit(1) & "/" & propSplit(2)
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" - "&arrStatus(0) &" performed successfully on "&arrObj(0)
						End If
						If  propSplit(1) < 10 Then 
							propSplit(1) = 0 & propSplit(1)
							Environment.Value(lcase(VarName)) = propSplit(0) & "/" & propSplit(1) & "/" & propSplit(2)
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" - "&arrStatus(0) &" performed successfully on "&arrObj(0)
						End If		
					Case "d"
						Environment.Value(lcase(VarName)) = FormatDateTime(Now(),2)
						Environment.Value(lcase(VarName)) = DateAdd("d",strstatus(1),Cdate(Environment.Value(lcase(VarName))))
						propSplit = Split(Environment.Value(lcase(VarName)),"/",-1,1)
						If  propSplit(0) < 10 Then
							propSplit(0) = 0 & propSplit(0)
							Environment.Value(lcase(VarName)) = propSplit(0) & "/" & propSplit(1) & "/" & propSplit(2)
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" - "&arrStatus(0) &" performed successfully on "&arrObj(0)
						End If
						If  propSplit(1) < 10 Then 
							propSplit(1) = 0 & propSplit(1)
							Environment.Value(lcase(VarName)) = propSplit(0) & "/" & propSplit(1) & "/" & propSplit(2)
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" - "&arrStatus(0) &" performed successfully on "&arrObj(0)
						End If	
					Case  "m"
						Environment.Value(lcase(VarName)) = FormatDateTime(Now(),2)
						Environment.Value(lcase(VarName)) = DateAdd("m",strstatus(1),Cdate(Environment.Value(lcase(VarName))))
						propSplit = Split(Environment.Value(lcase(VarName)),"/",-1,1)
						If  propSplit(0) < 10 Then
							propSplit(0) = 0 & propSplit(0)
							Environment.Value(lcase(VarName)) = propSplit(0) & "/" & propSplit(1) & "/" & propSplit(2)
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" - "&arrStatus(0) &" performed successfully on "&arrObj(0)
						End If
						If  propSplit(1) < 10 Then 
							propSplit(1) = 0 & propSplit(1)
							Environment.Value(lcase(VarName)) = propSplit(0) & "/" & propSplit(1) & "/" & propSplit(2)
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" - "&arrStatus(0) &" performed successfully on "&arrObj(0)
						End If	
					Case  "y"
						Environment.Value(lcase(VarName)) = FormatDateTime(Now(),2)
						Environment.Value(lcase(VarName)) = DateAdd("yyyy",strstatus(1),Cdate(lcase(Environment.Value(VarName))))
						propSplit = Split(Environment.Value(lcase(VarName)),"/",-1,1)
						If  propSplit(0) < 10 Then
							propSplit(0) = 0 & propSplit(0)
							Environment.Value(lcase(VarName)) = propSplit(0) & "/" & propSplit(1) & "/" & propSplit(2)
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" - "&arrStatus(0) &" performed successfully on "&arrObj(0)
						End If
						If  propSplit(1) < 10 Then 
							propSplit(1) = 0 & propSplit(1)
							Environment.Value(lcase(VarName)) = propSplit(0) & "/" & propSplit(1) & "/" & propSplit(2)
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" - "&arrStatus(0) &" performed successfully on "&arrObj(0)
						End If
				End Select
				arrKeyIndex(1) = Environment.Value(lcase(VarName))
			End If
			arrKeyIndex(1) = GetValue(arrKeyIndex(1))
			If (LCase(arrObj(0)) = "textbox") Or (LCase(arrObj(0)) = "webfile") Or (LCase(arrObj(0)) = "tabletextbox") Then
				object.Set (arrKeyIndex(1))
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			ElseIf (LCase(arrObj(0)) = "checkbox") Or (LCase(arrObj(0)) = "tablecheckbox") Then
				object.Set arrKeyIndex(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			ElseIf (LCase(arrObj(0))="element") Then 'customization
				object.Click
				Wait(2)
				Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys arrKeyIndex(1)
				Set WshShell=Nothing
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) & ". Please verify Keyword entered."
			End If
		Case "click"
				object.Click
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		Case "press"
				object.press arrKeyIndex(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		Case "select"
			If LCase(arrObj(0)) = "radiobutton" Or LCase(arrObj(0)) = "combobox"  Or LCase(arrObj(0)) = "tablecombobox" Or LCase(arrObj (0)) = "tableradiobutton" Or LCase(arrObj (0)) = "menu" Then
			object.Select arrKeyIndex(1)
  				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) & ". Please verify Keyword entered."
			End If
		Case "selectindex"
			If LCase(arrObj(0)) = "radiobutton" Or LCase(arrObj(0)) = "combobox"  Or LCase(arrObj(0)) = "tablecombobox" Or LCase(arrObj (0)) = "tableradiobutton" Then
				object.Select("#"&arrKeyIndex(1))
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			End If
	Case "verifyselect"
			Dim flag  			'Stores the flag value
			Dim actualItem		'Stores the value to be checked
			Dim actualListCount	'Stores the total number of elements in the listBox
			Dim loopCurrentItem	'used for looping
			flag = 2
			actualListCount =object.GetROProperty("items count")
			For loopCurrentItem = 1 To (actualListCount)
				actualItem = Cstr(object.GetItem(loopCurrentItem ))
				'End the Loop if a match is found
				If  Instr(actualItem,arrKeyIndex(1)) <> 0 Then
					flag = 1
				  Exit For
				End If
			Next
				If flag = 1 Then
					object.Select(actualItem)
					Reporter.reportevent micPass, arrKeyValue(1) & "should be selected", "Expected Item " & arrKeyValue(1) & "is selected." 
				Else
				    Reporter.reportevent micFail, arrKeyValue(1) & "should be selected", "Expected Item " & arrKeyValue(1) & "not present in the list, hence select operation cannot be performed, Check line number" & intRowCount
				End If
		Case "close"
			If LCase(arrObj(0)) = "browser" Or LCase(arrObj(0)) = "window" Or LCase(arrObj(0)) = "dialog" Then
				parent.Close
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				 Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) & "'. Please verify Keyword entered."
			End If
			Case "restore"
				If  LCase(arrObj(0)) = "dialog" Then
				parent.Restore
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				 Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) & "'. Please verify Keyword entered."
			End If
			Case "activate"
				If  LCase(arrObj(0)) = "dialog" Then		
				parent.Activate	
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				 Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) & "'. Please verify Keyword entered."
			End If
			Case "textclick"
				If  LCase(arrObj(0)) = "dialog" Then		
				Call Func_SelectText(arrKeyIndex(1))	
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				 Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) & "'. Please verify Keyword entered."
			End If
		Case "submit"
			object.Submit  'The object should be focused in this case
			Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		Case "fire event"
			If LCase(arrObj(0)) <> "webfile" Then
				object.FireEvent arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) & "'. Please verify Keyword entered."
			End If
		Case "deselect"
			If LCase(arrObj(0)) = "combobox"  Or LCase(arrObj(0)) = "tablecombobox" Then
				object.Deselect arrKeyIndex(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
			   Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) & "'. Please verify Keyword entered."
			End If
		Case "extendselect"
			If LCase(arrObj(0)) = "combobox" Or LCase(arrObj(0)) = "tablecombobox" Then
				object.ExtendSelect arrKeyIndex(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
			  Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) & "'. Please verify Keyword entered."
			End If
		Case Else
			If (LCase(Trim(arrObj(0)))= "sqlexecute") Or (LCase(Trim(arrObj(0)))= "sqlvaluecapture") Or (LCase(Trim(arrObj(0)))= "sqlcheckpoint") Then
'				NOTE: The following variables should be present in the Environment_values.xml File attached to the script
				Environment("connectionString") = "DRIVER={" & Environment("dbDRIVER") & "};UID=" & Environment("dbUID") & ";PWD=" & Environment("dbPWD") & ";SERVER=" & Environment("dbServer") & "_" & Environment("dbHost") & ".world"
				Select Case LCase(Trim(arrObj(0))) 'To execute SQL operations
					Case "sqlexecute"
						strSQL = arrObj(1)
						For inti = 0 to (Func_RegExpMatch ("##\w###",arrObj(1),aPosition,aMatch) - 1)
							strReplace = "'" & Environment.Value(Replace(aMatch(inti),"##","",1,-1,1)) & "'"
							strSQL = Func_gfRegExpReplace(aMatch(inti), strSQL, strReplace)
						Next								
						Set dbConn = CreateObject("ADODB.Connection")	'Create the DB Object
						dbConn.Open Environment.Value("connectionString")
						Set dbRs = dbConn.Execute(strSQL)	'Execute the query
						dbConn.Close	'Close the database connection
						Set dbConn = Nothing
						Reporter.ReportEvent micDone, " Sql Operation", "Query executed successfully" 		
					Case "sqlvaluecapture"
						Environment.Value(lcase(arrKeyIndex(0))) = Func_gfQuery(arrObj(1))
						Reporter.ReportEvent micDone, " Sql value capture", "Sql value captured successfully" 
					Case "sqlcheckpoint"
						DataTable.GetSheet("Action1").SetCurrentRow(1)
						DbTable(arrKeyIndex(0)).SetTOProperty "connectionstring", Environment.Value("connectionString")	'Set the TO property of connection string.
						strSQL = arrAction(1)
						For inti = 0 to (Func_RegExpMatch ("##\w##", arrAction(1),aPosition,aMatch) - 1)
							strReplace = "'" & Environment.Value(Replace(aMatch(inti),"##","",1,-1,1)) & "'"
							strSQL = Func_gfRegExpReplace(aMatch(inti), strSQL, strReplace)
						Next
						DbTable(arrKeyIndex(0)).SetTOProperty "source", strSQL	'Change the DB OBjects source(SQL) statement.
						DbTable(arrKeyIndex(0)).Check CheckPoint(arrKeyIndex(0)) 'Execute the DB checkpoint.
						Reporter.ReportEvent micDone, " sql check point", "sql check point successfully" 
				End Select
			Else
				If (LCase((arrObj(0)) <> "split") And (LCase(arrObj(0)) <> "random")) Then
					errStr1 = "Keyword Check at Line no - " & intRowCount 
					errStr2 = "Keyword - '" & arrkeyValue(0) & "'  not supported. Please verify Keyword entered."
					err.raise vbObjectError + vbObjectError
				End If	
			End If 
	End Select 
	Dim intNum      'This variable stores the value in arrObj(1)
	Dim strvar      'This variable is used to store the string value in 4th Column of Datasheet
	Dim strsplit    'This variable is used to store the array after Split operation
	Dim strlen      'This variable is used to store the string value in 5th Column of Datasheet
	Dim strstore1    'This variable is used to store the elements present in the fourth Column 
	Dim arrVals		'This variable is used store the split array elements
	Dim strvarstore 'This variable is used to store the split element 	
	Dim intval
	Select Case LCase(arrObj(0)) 'To execute special/common operations
		Case "random"
			intNum = arrObj(1)
			  strvar = Cstr(Trim(DataTable.Value(4,dtLocalSheet)))
			Environment.Value(lcase(strvar)) = Rnd(intNum)
		Case "split"
			strvar = Split(arrObj(1),"^",-1,1)
			For inti= 0 to Ubound(strvar)
				If(Instr(1,strvar(inti),"#") = 1) Then
					strvar(inti) = Environment.Value(lcase(Right(strvar(inti),Len(strvar(inti))-1)))
				End If
			Next
			strsplit = Split(strvar(0),strvar(1),-1,1)
			If  DataTable.Value(5,dtLocalSheet) <> "" Then
			   strlen=Cstr(Trim(DataTable.Value(5,dtLocalSheet)))
				 Environment.Value(lcase(strlen))=Ubound(strsplit)
			End If
			strstore1 = Cstr(Trim(DataTable.Value(4,dtLocalSheet)))
			arrVals = Split(strstore1,";",-1,1)
			For inti = 0 to Ubound(arrVals)
				strvarstore  = Split(arrVals(inti),":",2,1)
				intval = Cint(strvarstore(1))
				Environment.Value(lcase(strvarstore(0))) = strsplit(intval)
			Next			   
	End Select 'End of special/common operations  
		If   Environment.Value ("iperform") =1 Then
		Call Func_CaptureScreenshot("test",intRowCount)  'call screencapture function to take screenshot
	End If
End Function 'End of perform
'#################################################################################################
'#################################################################################################
'Function name 		: Func_Store_Web
'Description    	: If the user requires to store any property of a particular object into a variable 
'					then this function can be used.
'Parameters     	: The Object details as the full hierarchical Path of the Object goes as parameter 
'					to the function.            
'Assumptions    	: None 
'#################################################################################################
'The following function is for StoreValue keyword.
'#################################################################################################
Function Func_Store_Web(object,arrObj)
	Dim strPropName  'Stores the name of the Property to be stored
	Dim arrPropSplit 'Stores the Property name and Variable Name.
	Dim intGRowNum	'Stores the row number
	Dim intGColNum	'Stores the Column number
	arrPropSplit = Split(Datatable.Value(4,dtLocalSheet),":",-1,1) 'splitting the value in the 4th Column into Property and Variable Name
	strPropName = arrPropSplit(0)		  'Storing the Property into strPropName 
	VarName = arrPropSplit(1)			  'Storing the Variable Name into VarName 
	Select Case LCase(strPropName)		  'Case to store the required property of the variable into the VarName variable 
		Case "itemscount"
			Environment.Value(lcase(VarName)) = object.GetROProperty("items count")
		Case "enabled"
			Environment.Value(lcase(VarName)) = Not(CBOOL(object.GetROProperty("disabled")))
		Case "columncount"
			Environment.Value(lcase(VarName)) = object.ColumnCount(1)
		Case "rowcount"
			Environment.Value(lcase(VarName)) = object.RowCount
		Case "filename"
			Environment.Value(lcase(VarName)) = object.GetROProperty("file name")
		Case "imagetype"
			Environment.Value(lcase(VarName)) = object.GetROProperty("image type")
		Case "defaultvalue"
			Environment.Value(lcase(VarName)) = object.GetROProperty("default value")
		Case "maxlength"
			Environment.Value(lcase(VarName)) = object.GetROProperty("max length")
		Case "allitems"
			Environment.Value(lcase(VarName)) = object.GetROProperty("all items")
		Case "selectiontype"
			Environment.Value(lcase(VarName)) = object.GetROProperty("select type")
		Case "exist"
			If arrObj(0) = "window" Or arrObj(0) = "dialog" Or arrObj(0) = "browser" Then
				Environment.Value(lcase(VarName)) = Cstr((curParent.Exist(5)))
			Else
				Environment.Value(lcase(VarName)) = Cstr((object.Exist(5)))
			End If
		Case "selectioncount"
			Environment.Value(lcase(VarName)) = object.GetROProperty("selected items count")
		Case "getcelldata"
			If arrObj(0)="table" Then
				intGRowNum = Cint(arrPropSplit(2))
				intGColNum = Cint(arrPropSplit(3))
				Environment.Value(lcase(VarName)) = object.GetCellData(intGRowNum,intGColNum)
			Else
				Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & strPropName & "'  not supported for -" &arrObj(0) & ". Please verify Keyword entered."
			End If
		Case Else
			If arrObj(0) = "window" Or arrObj(0) = "dialog" Or arrObj(0) = "browser" Then
				Environment.Value(lcase(VarName)) = curParent.GetROProperty(strPropName)
			Else
				Environment.Value(lcase(VarName)) = object.GetROProperty(strPropName)
			End If
	End Select
    If cint(Introwcount)<=cint(Environment("intEndRow")) Then
		If Cint(Environment("intStartRow"))<=Cint(Introwcount) Then
			Call DebugGetEnv()   'to call debug function for execution log status in HTML file
        End If 
    End If
End Function
'#################################################################################################
'#################################################################################################
'Function name 	: Func_Check_Web
'Description    : This function is used for all the checking operations to be performed on the AUT.
'Parameters     : The Object details on which check has to be performed along with details of the 
'				fourth Column of the current row in Local sheet and the current row number in the 
'				Local Sheet.           
'Assumptions    : NA
'#################################################################################################
'The following function is for 'Check' keyword.
'#################################################################################################
Function Func_Check_Web(object,arrObj,arrKeyValue,arrKeyIndex,intRowCount)
	Dim propertyVal     'Stores the Property name which needs to be checked
	Dim strChecking     'Stores the property value retrieved from the AUT
	Dim checkval			'Stores the Pass or Fail response
	checkval = 2
	If Ubound(arrKeyValue)=0 Then
		arrKeyValue(1) = ""
	End If	

	Select Case LCase(arrKeyValue(0)) 'checking the property value of the object
           Case "tablesearch"
				searchtablereturn= Func_tablesearch(object,arrTableindex(1))
				reportStep=	"Verify that a row value matches for search criteria '"&arrTableindex(1)
				arrsearchtablereturn=Split(searchtablereturn,":")
				If Lcase(arrsearchtablereturn(1)) =Lcase( arrsearchtablereturn(2))Then
						reportStepPass="The row value corresponding for the search criteria " & arrTableindex(1)& "  is exists, which is as expected."
						Reporter.ReportEvent micPass,reportStep,reportStepPass
						 keyword=0
						checkval=1
				Elseif 	Lcase(arrsearchtablereturn(1)) <> Lcase(arrsearchtablereturn(2)) Then
						keyword=1
						checkval=0
						reportStepFail =  "The corresponding row values doesn't suit the search criteria '"& arrTableindex(1) &",  Which is  not as expected."
                		Reporter.ReportEvent  micFail,reportStep,reportStepFail
				End If
		Case "enabled"
			propertyVal = "disabled"
		Case "url"
			propertyVal = "url"   
		Case "search" 
			Dim strText
			Dim strSearch
			Dim strItem
			Dim strfnd
			Dim strf
			propertyVal = "null"
            strText=object.GetROProperty("items count")
            If strText<>0 Then
				For strf=1 to strText
					strItem= object.GetItem(strf)
					If Trim(strItem)= Trim(arrKeyIndex(1)) Then
						strfnd = 1
						Exit For
					End If
				Next
			Else
			   strfnd = 0
            End If
			If strfnd <> 0 Then
				Reporter.ReportEvent micPass,"Verify that '" & arrKeyIndex(1)&"'  Text should be present in '" & arrObj(0) & "' - " & arrObj(1),"'" & arrKeyIndex(1)&"'  Text is  present in '" & arrObj(0) & "' - " & arrObj(1)
				keyword=0
				checkval = 1
				If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	 ' calling function update log to create an execution log in HTML file
				End If
			Else
				Reporter.ReportEvent micFail,"Verify that '" & arrKeyIndex(1)&"'  Text should be present in '" & arrObj(0) & "' - " & arrObj(1),"'" & arrKeyIndex(1)&"'  Text is not present in '" & arrObj(0) & "' - " & arrObj(1)
				keyword=1
				checkval = 0
				Call Func_CaptureScreenshot("checkpoint",intRowCount)	 'call screencapture function to take screenshot 
				If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")' calling function update log to create an execution log in HTML file
				End If
			End If
		Case "visible"
			propertyVal = "visible"
		Case "focused"
			propertyVal = "focused"
		Case "selection"
			If arrObj(0) = "combobox" Or arrObj(0) = "tablecombobox" Then
				propertyVal = "selection"
			ElseIf arrObj(0) = "radiobutton" Or arrObj(0) = "tableradiobutton" Then
				propertyVal = "value"	
			Else
				Reporter.Reportevent micFail,"Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" & arrObj(0)& ". Please verify Keyword entered."
			End If

		Case "checked"
			If arrObj(0) = "checkbox" Or arrObj(0) = "tablecheckbox" Then
				propertyVal = "checked"				
			Else
				Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If

		Case "text"
			Dim strText1  	 'Stores the text returned by the AUT
			Dim intRow 		 'Stores the Row count of the Table
			Dim intCol 		 'Stores the Column Count of the particular row
			If arrObj(0) = "table" Or arrObj(0) = "childtable" Then 'Checking if object type is table or child table
				propertyVal = "null"
				strCellData = arrKeyIndex(1)
				Call GetValue(strCellData)
				intRow = parent.WebTable(arrObj(1)).RowCount
				For intj = 1 To intRow
					intCol = parent.WebTable(arrObj(1)).ColumnCount(intj)
					For inti = 1 To intCol                                                 
						strText1 = Trim(parent.WebTable(arrObj(1)).GetCellData(intj,inti))
						If strText1 = Trim(strCellData) Then   'Comparing the text between actual and expected values
							keyword = 0
							Exit For
						Else
							keyword = 1
						End If
					Next
					If keyword = 0 Then
						Exit For
					End If
				Next
				If keyword = 0 Then
					If Ubound(arrKeyIndex) >1 Then
						If Lcase(arrKeyIndex(2)) = "notexist" Then
							If FrameName = "" Then
								Reporter.reportevent micFail, "Verify that In Table:'"   & arrObj(1)  & vbcr &strCellData  & " field should not be present",  "In Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strCellData & "'field is  present"
								keyword=1
								checkval = 0
								Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'call screencapture function to take screenshot
								If htmlreport = "1" Then
								 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	 'calling function update log to create an execution log in HTML file
								End If
							Else
								Reporter.reportevent micFail, "Verify that under Screen - '"& FrameName & "' And  in Table:'"   & arrObj(1)  & vbcr &strCellData  & "' field should not be present",  "Under Screen - ' "& FrameName & "'' And  in Table:'"   & arrObj(1)  & "'" & vbcr & ", '"  & strCellData & "' field is  present"
								keyword=1
								checkval = 0
								Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'call screencapture function to take screenshot
						If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
					End If
							End If
						Else
							If FrameName = "" Then
								Reporter.ReportEvent micPass, "Verify that In Table:'"   & arrObj(1)  & vbcr & strCellData  &" field should be present","  In Table:'"   & arrObj(1)  & "'" & vbcr & ", '"  & Text &"' field is present at " &vbcr& "Row : " &intj &vbcr&"Column : "&intii
								keyword=0
								checkval=1
						If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
							Else
							   Reporter.ReportEvent micPass, "Verify that under Screen - '"& FrameName & "' And  in Table:'"   & arrObj(1)  & vbcr & strCellData  &" field should be present", "Under Screen - ' "& FrameName & "'' And  in Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strText1 &"' field is present at " &vbcr& "Row : " &intj &vbcr&"Column : "&inti
							   keyword=0
							   checkval=1
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	'calling function update log to create an execution log in HTML file
					End If
							End If	
						End If
					Else
						If FrameName = "" Then
							Reporter.ReportEvent micPass, "Verify that In Table:'"   & arrObj(1)  & vbcr & strCellData  &" field should be present","  In Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strText1 &"' field is present at " &vbcr& "Row : " &intj &vbcr&"Column : "&inti
							keyword=0
							checkval=1
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
						Else
							Reporter.ReportEvent micPass, "Verify that under Screen - '"& FrameName & "' And  in Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strCellData  &" field should be present", "Under Screen - ' "& FrameName & "' And  in Table:'"   & arrObj(1) & "'" & vbcr & ", '" & strText1 &"' field is present at " &vbcr& "Row : " &intj &vbcr&"Column : "&inti
							keyword=0
							checkval=1
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	 ' calling function update log to create an execution log in HTML file
					End If
						End If
					End If   
				Else
					If Ubound(arrKeyIndex) >1 Then
						If Lcase(arrKeyIndex(2)) = "notexist" Then
							If FrameName = "" Then
								Reporter.reportevent micPass, "Verify that In Table:'"   & arrObj(1)  & vbcr &strCellData  & " field should not be present",  "In Table:'"   & arrObj(1) & "'" & vbcr & ", '" & strCellData & "' field is not present"
								If htmlreport = "1" Then
								 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	'calling function update log to create an execution log in HTML file
								End If
								keyword=0
								checkval=1
							Else
								Reporter.reportevent micPass, "Verify that under Screen - '"& FrameName & "' And  in Table:'"   & arrObj(1)  & "'" & vbcr & ", '" &strCellData  & "' field should not be present",  "Under Screen - ' "& FrameName & "' And  in Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strCellData & "' field is not present"
								keyword=0
								checkval=1
								If htmlreport = "1" Then
									 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	 ' calling function update log to create an execution log in HTML file
								End If
							End If
						Else
							If FrameName = "" Then
								Reporter.ReportEvent micFail, "Verify that In Table:'"   & arrObj(1)  & "'" & vbcr & ", '"& strCellData  &" field should be present"," In Table:'"   & arrObj(1)  & "'" & vbcr & ", '"  & strCellData & "' field is not present"
								keyword=1
								checkval=0
								Call Func_CaptureScreenshot("checkpoint",intRowCount)	 'call  to screencapture function to take screenshot
								If htmlreport = "1" Then
								 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	 ' calling function update log to create an execution log in HTML file
								End If
							Else
								Reporter.ReportEvent micFail, "Verify that under Screen - '"& FrameName & "' And  in Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strCellData  &"' field should be present",  "Under Screen - ' "& FrameName & "' And  in Table:'"   & arrObj(1) & "'" & vbcr & ", '" & strCellData & "' field is not present"
								keyword=1
								checkval = 0
								Call Func_CaptureScreenshot("checkpoint",intRowCount)	   'call  to screencapture function to take screenshot
						If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
					End If
							End If	
						End If
					Else
						If FrameName = "" Then
							Reporter.ReportEvent micFail, "Verify that In Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strCellData  &"' field should be present"," In Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strCellData & "' field is not present"
							keyword=1
							checkval=0
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	''call to screencapture function to take screenshot
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	 ' calling function update log to create an execution log in HTML file
					End If
						Else
							Reporter.ReportEvent micFail, "Verify that under Screen - '"& FrameName & "' And  in Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strCellData  &"' field should be present",  "Under Screen - ' "& FrameName & "' And  in Table:'"   & arrObj(1)  & "'" & vbcr & ", '" & strCellData & "' field is not present"
							keyword=1
							checkval=0
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	'call to screencapture function to take screenshot
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	 ' calling function update log to create an execution log in HTML file
					End If
						End If	
					End If	
				 End If
			ElseIf arrObj(0) = "element" Or arrObj(0) = "tableelement" Then
				propertyVal = "innertext"
			ElseIf arrObj(0) = "textbox" Or arrObj(0) = "tabletextbox" Then
				propertyVal = "value"
			Else	
				propertyVal = "text"
			End If

		Case "itemscount"
			If arrObj(0) = "combobox" Or arrObj(0) = "radiobutton"  Or arrObj(0) = "tablecombobox" Or arrObj(0) = "tableradiobutton"Then
				propertyVal = "items count"
			Else
				Reporter.ReportEvent micFail,  "Keyword not supported by Framework", "This operation cannot be performed"
			End If

		Case "exist"
			 propertyVal = "null"
				If arrObj(0) = "browser" Or arrObj(0) = "window" Or arrObj(0) = "dialog"  Or arrObj(0) = "frame" Or arrObj(0) = "page" Then
					If curParent.exist Then
						If Lcase(arrKeyIndex(1)) = "true"  Then
							Reporter.reportevent micPass, "Verify that " & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  exist","The " & arrObj(0) & "-"&"'" & arrObj(1) &"'"& " exists which is as expected"
							checkval=1
							keyword = 0
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
						ElseIf Lcase(arrKeyIndex(1)) = "false" Then
							Reporter.reportevent micFail, "Verify that " & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  not exist","The " & arrObj(0) & "-"&"'" & arrObj(1) &"'"& "  exists which is  not as expected"
							keyword = 1
							checkval=0
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'call  to screencapture function to take screenshot
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
					End If
						End If
					Else
						If Lcase(arrKeyIndex(1)) = "false" Then
							 Reporter.reportevent micPass, "Verify that " & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  not exist","The " &  arrObj(0) & "-"&"'" & arrObj(1) &"'"& " does not exist which is as expected"
							keyword = 0
							checkval=1
							If htmlreport = "1" Then
								Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
							End If	
						ElseIf Lcase(arrKeyIndex(1)) = "true"  Then
						   Reporter.reportevent micFail, "Verify that " &  arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  exist","The " & arrObj(0) & "-"&"'" & arrObj(1) &"'"& " does not exist which is  not as expected"
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
					End If
							keyword = 1
							checkval=0 
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	 'call  to screencapture function to take screenshot
						End If
					End If
				ElseIf (arrObj(0) = "tablecombobox" Or arrObj(0) = "tableimage" Or arrObj(0) = "tablebutton" Or arrObj(0) = "tablelink" Or arrObj(0) = "tablecheckbox" Or arrObj(0) = "tabletextbox" Or arrObj(0) ="tableradiobutton") or arrObj(0) = "tableelement" or arrObj(0) = "childtable"   Then
					Select Case Lcase (arrObj(0)) 'case statement for table operations
						Case "tabletextbox"
							  ObjectVal = "WebEdit"
						Case "tablebutton"
							ObjectVal = "WebButton"
						Case "tablecombobox"
							ObjectVal = "WebList"
						Case "tablecheckbox"
							ObjectVal = "WebCheckBox"
						 Case "tableradiobutton"
							 ObjectVal ="WebRadioGroup"
						 Case "tablelink"
							 ObjectVal = "Link"
						Case "tableelement"
							ObjectVal ="WebElement"
						Case "childtable"
							ObjectVal = "WebTable"
						Case "tableimage"
							  ObjectVal = "Image"
						Case Else
							errStr1 = "Keyword not supported by Framework at Line no.- " & intRowCount
							errStr2 = "This operation cannot be performed"
							err.raise vbObjectError
					End Select
					If (Instr(1,arrKeyIndex(1),"#") <> 0) Then
						arrKeyIndex(1) = Environment.Value(Right(arrKeyIndex(1),Len(arrKeyIndex(1))-1))
					End If
					 If (Instr(1,arrKeyIndex(2),"#") <> 0) Then
						arrKeyIndex(2) = Environment.Value(Right(arrKeyIndex(2),Len(arrKeyIndex(2))-1))
					End If	
					childCount = parent.WebTable(arrObj(1)).childItemCount(CInt(arrKeyIndex(1)), CInt(arrKeyIndex(2)),ObjectVal)	
					If childCount <> 0 Then
						If FrameName = "" Then
							Reporter.reportevent micPass, "Verify that  in table" & "'" & arrObj(1) & "','" & arrObj(0) & "'  should  exist in Row:" & arrKeyIndex(1)& vbcr & "Column:" & arrKeyIndex(2),  "In table"&"'" & arrObj(1) &"','" &arrObj(0) & "', exists  in Row:" & arrKeyIndex(1)& vbcr & "Column:" & arrKeyIndex(2) & " ,which is as expected"
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
                        Else 
							Reporter.reportevent micPass,"Verify that under Screen - ' "& FrameName & "'"  &vbcr  & "in table" & "'" & arrObj(1) & "','" & arrObj(0) & "'  should  exist in Row:" & arrKeyIndex(1)& vbcr & "Column:" & arrKeyIndex(2), "Under Screen- ''"& FrameName & "''," & vbcr &  "In table"&"'" & arrObj(1) &"','" &arrObj(0) & "', exists  in Row:" & arrKeyIndex(1)& vbcr & "Column:" & arrKeyIndex(2) & " , which is as expected"  						
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	  'calling function update log to create an execution log in HTML file
					End If
                        End If
						keyword = 0
						checkval=1
					Else
						If FrameName = "" Then
							Reporter.reportevent micFail,"Verify that " & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  exist", arrObj(0) & "-"&"'" & arrObj(1) &"'"& " does not exist which is  not as expected"
							keyword=1
							checkval=0
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'Call to capture screenshot function
						If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
					End If
						Else
							Reporter.reportevent micFail, "Verify that under Screen -'" & FrameName &"'" & vbcr & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  exist", "Under Screen - '" & FrameName & "'," & vbcr & arrObj(0) & "-"&"'" & arrObj(1) &"'"& " does not exist  which is  not as expected"					 
							keyword=1
							checkval=0
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	  ' Call to capture screenshot function
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
					End If
						End If
        			End If    
				Else
				If object.exist Then
					If Lcase(arrKeyIndex(1)) = "true" Then
						If FrameName = "" Then
							Reporter.reportevent micPass, "Verify that " & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  exist", arrObj(0) & "-"&"'" & arrObj(1) &"'"& " exists, which is as expected"
							keyword=0
							checkval=1
						If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
						Else 
							Reporter.reportevent micPass,"Verify that under Screen - ' "& FrameName & "'"  &vbcr & arrObj(0) & "-" & "'" & arrObj(1) & "' should  exist", "Under Screen- '"& FrameName & "'" & vbcr & arrObj(0) & "-"&"'" & arrObj(1) &"' exists, which is as expected"  						
							If htmlreport = "1" Then
								Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
							End If	
						End If
						keyword = 0
						checkval=1
					ElseIf Lcase(arrKeyIndex(1)) = "false" Then
						If FrameName = "" Then
							Reporter.reportevent micFail, "Verify that " &  arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should not exist", arrObj(0) & "-"&"'" & arrObj(1) &"'"& "  exists which is not as expected"
							keyword=1
							checkval=0
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	'Call to capture screenshot function
						If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
					End If
						Else
							Reporter.reportevent micFail, "Under Screen-'" & FrameName &"'" & vbcr & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should not exist", "Under Screen - '" & FrameName & "'" & vbcr & arrObj(0) & "-"&"'" & arrObj(1) &"'"& " exists, which is  not as expected"					 
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'Call to capture screenshot function
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	 ' calling function update log to create an execution log in HTML file
					End If
						End If
						keyword = 1
						checkval=0
					End If	
				Else	
					If Lcase (arrKeyIndex(1)) = "false" Then
						If FrameName = "" Then
							Reporter.reportevent micPass, "Verify that " & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  not exist", arrObj(0) & "-"&"'" & arrObj(1) &"'"& " does not exist, which is as expected"
							keyword=0
							checkval=1
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
						Else 
							Reporter.reportevent micPass,"Verify that under Screen - ' "& FrameName & "'"  &vbcr & arrObj(0) & "-" & "'" & arrObj(1) & "', should not  exist", "Under Screen- '"& FrameName & "'" & vbcr & arrObj(0) & "-"&"'" & arrObj(1) &"' does not exist, which is as expected"  						
						End If
						keyword = 0
						checkval=1
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
					ElseIf Lcase(arrKeyIndex(1)) = "true" Then
						If FrameName = "" Then
							Reporter.reportevent micFail, "Verify that " &  arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  exist", arrObj(0) & "-"&"'" & arrObj(1) &"'"& " does not exist which is  not as expected"
							keyword=1
							checkval=0
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'Call to capture screenshot function
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	 ' calling function update log to create an execution log in HTML file
					End If
						Else
							Reporter.reportevent micFail, "Verify that under Screen - '" & FrameName &"'" & vbcr & arrObj(0) & "-" & "'" & arrObj(1) & "'" & " should  exist", "Under Screen - '" & FrameName & "'," & vbcr & arrObj(0) & "-"&"'" & arrObj(1) &"'"& " does not exist, which is  not as expected"					 
							keyword=1
							checkval=0
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	  ' Call to capture screenshot function
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
					End If
						End If
						keyword = 1
						checkval=0
					End If
				End If
			End If
		End Select 'Property setting done. Table check/text check for table also done
		If (Lcase(arrKeyValue(0)) <> "tablesearch") Then
			If propertyVal <> "null" Then
			Select Case Lcase(arrKeyValue(0)) 'getting the run time properties of an object
				Case "enabled" 
					strChecking = CBool(object.getROProperty(propertyVal))
					strChecking=Not(strChecking)
				Case "checked"
                    strChecking = object.getROProperty(propertyVal)
					If Lcase(strChecking) = 1  Then
						strChecking = "True"
			        ElseIf Lcase (strChecking)= 0 Then
						strChecking = "False"  
					End If
				Case "focused"
					strChecking = object.getROProperty(propertyVal)
					If Lcase(strChecking) = 1  Then
						strChecking = "True"
			        ElseIf Lcase (strChecking)= 0 Then
						strChecking = "False"  
					End If
				Case Else
                    strChecking = object.getROProperty(propertyVal)
			End Select
			If Trim(CStr(Lcase(strChecking))) =Trim(Lcase(arrKeyIndex(1))) Then 
				If (arrKeyIndex(0)="tabselection") Then 
					Reporter.reportevent micPass,"Verify that  in " & arrObj(0) & "-" & "'" & arrObj(1) & "' , '" & "  is selected", arrObj(0) & "-"&"' " & arrObj(1) &"' , '"  & "  is selected, which is as expected"  
					keyword=0
					checkval=1
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
				ElseIf (arrKeyValue(0) = "text") Then
					If arrKeyIndex(1)= "" Then
					arrKeyIndex(1) = "blank"
				End If
				If strChecking = "" Then
					strChecking = "blank"
				End If
				Reporter.reportevent micPass,"Verify that text in " & arrObj(0) & "-" & "'" & arrObj(1) & "'"& " is " & arrKeyIndex(1)  ,"Text in " &arrObj(0) & "-"&"'" & arrObj(1) &"' "& " is " & strChecking & ", which is as expected"
				keyword=0
				checkval=1
				If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
				End If
			ElseIf (arrKeyValue(0) = "bgcolor") Then
				Reporter.reportevent micPass,"Verify that " & arrKeyIndex(0) &  " of  " & arrObj(0) & "-" & "'" & arrObj(1) & "'"& " is " & arrKeyIndex(1)  ,arrKeyIndex(0) & " of " & arrObj(0) & "-"&"'" & arrObj(1) &"' "& " is " & arrKeyIndex(1) & " which is as expected" 	 
				keyword=0
				checkval=1
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
			Else				
				Reporter.reportevent micPass,"Verify that in " & arrObj(0) & "-" & "'" & arrObj(1) & "'"& ", Property '" & arrKeyIndex(0) & "' is " & arrKeyIndex(1)  ,"In " & arrObj(0) & "-"&"'" & arrObj(1) &"' "& ", Property '" & arrKeyIndex(0)& "' is " & strChecking & ", which is as expected" 
				keyword=0
				checkval=1
				If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If
			End If

			Else
				If LCase((arrKeyIndex(0))="tabselection") Then 
					Reporter.reportevent micFail,"Verify that  in " & arrObj(0) & "-" & "'" & arrObj(1) & "' , '" & "  is selected", arrObj(0) & "-"&"' " & arrObj(1) &"' , '"  & "  is not selected, which is not as expected"  
					keyword=1
					checkval=0
					Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'Call to capture screenshot function
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
					End If
				ElseIf (arrKeyValue(0) = "text") Then
					If arrKeyIndex(1)= "" Then
					arrKeyIndex(1) = "blank"
				End If
				If strChecking = "" Then
					strChecking = "blank"
				End If
				Reporter.reportevent micFail,"Verify that text in " & arrObj(0) & "-" & "'" & arrObj(1) & "'"& " is " & arrKeyIndex(1)  ,"Text in " &arrObj(0) & "-"&"'" & arrObj(1) &"' "& " is " & strChecking & ", which is not as expected"
				keyword=1
				checkval=0
				Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'Call to capture screenshot function
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	
					End If
			ElseIf (arrKeyValue(0) = "bgcolor") Then
				Reporter.reportevent micFail,"Verify that " & arrKeyIndex(0) &  " of  " & arrObj(0) & "-" & "'" & arrObj(1) & "'"& " is " & arrKeyIndex(1)  ,arrKeyIndex(0) & " of " & arrObj(0) & "-"&"'" & arrObj(1) &"' "& " is " & strChecking & " which is not as expected" 	 
				keyword=1
				checkval=0
				Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'Call to capture screenshot function	
				If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
				End If
			Else				
				Reporter.reportevent micFail,"Verify that in " & arrObj(0) & "-" & "'" & arrObj(1) & "'"& ", Property '" & arrKeyIndex(0) & "' is " & arrKeyIndex(1)  ,"In " & arrObj(0) & "-"&"'" & arrObj(1) &"' "& ", Property '" & arrKeyIndex(0)& "' is " & strChecking & ", which is not as expected" 
				keyword = 1
				checkval=0
				If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	' calling function update log to create an execution log in HTML file
				End If
				Call Func_CaptureScreenshot("checkpoint",intRowCount)	  'Call to capture screenshot function
			End If
			End If	
		 End If	
	End If

    If Datatable.Value(5,dtLocalSheet) <> empty Then
		Call Func_CheckValueReturn(intRowCount,checkval) 
						 If cint(Introwcount)<=cint(Environment("intEndRow")) Then
			    	If Cint(Environment("intStartRow"))<=Cint(Introwcount) Then
				   Call DebugGetEnv() 		 'to call debug function for execution log status in HTML file
				   End If	
				End If
	End If

		If   Environment.Value("icheck")=1  Then
			Call Func_CaptureScreenshot("test",intRowCount)   'Call to capture screenshot function	
		End If

End Function
'#######################################################################################################
'Function name 	: Func_getRowNum_Web
'Description        : If the user wants to retrieve the row number in which specified text is present in table, this function can be used
'Parameters       	: The object(i.e, Table)  in which the search operation needs to be performed
'					 				 The text to be searched in the table and the number of columns in the table.
'Return Value		: This function returns row number of the given celltext
'Assumptions     	: NA
'#####################################################################################################
'The following function is used for rownum Keyword
'#####################################################################################################
Function Func_getRowNum_Web(object,strSearch,strReturnVal)
	Dim arrCol			'Stores the number of columns	
	Dim intCheck  'stores the row number
	Dim intStart   'initialization variable
	Dim strRowVal 'stores row value defined in datatable
	Dim strReturnVal1  'stores the return value of row number from this function
	initial = DataTable.Value(2,dtLocalSheet)
	strReturnVal1=DataTable.Value(5,dtLocalSheet)
	intCheck = -1
	arrCol = Split(strSearch,"--")
	If (instr(1,arrCol(0),"#")>0) Then
	arrCol(0)= Environment.Value(lcase(Right(arrCol(0),Len(arrCol(0))-1)))
	End If
	If (Ubound(arrCol) > 0) Then
	   intStart = Cint(arrCol(1))
	Else
		intStart = 1
	End If
	strRowVal = arrCol(0)
	intCheck = object.GetRowWithCellText(strRowVal,,intStart)

		If (intCheck>-1 ) Then
  		 Reporter.ReportEvent micPass,"Cell text '"&strRowVal&"'  should  be present in the table","Cell text '" &strRowVal& "' is  present in '" &intCheck &"' row of  the table"
      	Else
		Reporter.ReportEvent micFail,"Cell text "&strRowVal&"  should present in the table","Cell text '" &strRowVal& "' is  not  present in  the table"
    	End If
    Environment(strReturnVal1) = intCheck
End Function
'#####################################################################################################
'#####################################################################################################
'Function name 	: Func_DragDrop_Web
'Description        : Drags and Drop the object from the original position to the specified position
'Parameters       	: 		arrKeyValue:value from 4th column of datatable 
'										Object	 :: This is the object on which the specified operation needs to be performed.  .
'Assumptions     	: NA
'#####################################################################################################
Function Func_DragDrop_Web(arrKeyValue,object)
Dim deviceReplay 'Object created to access library functions
Dim ix							'Stores the property value of abs_x
Dim iy							'Stores the property value of abs_y
Dim fx							'Stores the value of arrKeyvalue(0)
Dim fy							'Stores the value of arrKeyvalue(1)
fx = arrKeyValue(0)
fy =  arrKeyValue(1)
Set deviceReplay = CreateObject( "Mercury.DeviceReplay")
ix = object.getROProperty("abs_x") 
iy = object.getROProperty("abs_y")
object.Activate
deviceReplay.DragAndDrop ix, iy, fx, fy, mb 			'Original  and drag position is mentioned
Reporter.ReportEvent micPass,"Drag - Drop", "Object has been moved to the specified position Successfully"
Set deviceReplay = Nothing
End Function
'#########################################################################################
'#########################################################################################
'Function name 	    : Func_Wait_Web
'Description        : This function is used for synchronization  with the application
'Parameters       	: The 'Object type' and the 'action being performed is passed as parameters. 
'Assumptions     	: None
'#########################################################################################
'The following function is used internally.
'#########################################################################################
Function Func_Wait_Web(arrObj,arrKeyValue,initial)
	On error resume next
	If Lcase(Trim(initial)) = "perform" or Lcase(Trim(initial)) = "context" Then
		   If(Ubound(arrKeyValue) >= 0)  Then
				If (LCase((arrKeyValue(0)) <> "exist" ) and (LCase(arrKeyValue(0)) <> "visible" )) Then
					curParent.WaitProperty "visible",True,20
				End If
				If (LCase((arrKeyValue(0)) <> "enabled") And LCase((arrKeyValue(0)) <>"exist") and LCase((arrKeyValue(0)) <> "visible" )) Then
					If LCase((arrObj(0)) = "button" Or LCase(arrObj(0)) = "checkbox" Or LCase(arrObj(0)) = "textbox" Or LCase(arrObj(0)) ="radiobutton" Or LCase(arrObj(0)) = "tablebutton" Or LCase(arrObj(0)) = "tablecheckbox" Or LCase(arrObj(0)) = "tabletextbox" Or LCase(arrObj(0)) ="tableradiobutton") Then
						object.WaitProperty "disabled",0,20
					ElseIf LCase(arrObj(0)) = "table" Or LCase(arrObj(0)) = "combobox" Or LCase(arrObj(0)) = "listbox" Or LCase(arrObj(0))= "element" Or LCase(arrObj(0)) = "link"  Or LCase(arrObj(0)) = "tablecombobox" Or LCase(arrObj(0))= "tableelement" Or LCase(arrObj(0)) = "tablelink"   Then
						object.WaitProperty "visible",True,20
					End If
				End If
			End If
		   End If 
End Function
'##########################################################################################
