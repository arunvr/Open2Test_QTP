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
	End If
	Set object = Nothing
	Call Func_ObjectSet_Win(arrAction,intRowCount)
	Select Case LCase(initial) ' to perform keyword operation defined in datatable
		Case "perform" 'To select the operations
			'start perform
			Call Func_Perform_Win(object,arrAction,arrKeyValue,arrKeyIndex,intRowCount)
        Case "storevalue"
			Call Func_Store_Win(object)
            'Checking
		Case "check"
			Call Func_Check_Win(object,arrAction,arrKeyValue,arrKeyIndex,intRowCount)
		Case "dragdrop"
			Call Func_DragDrop_Win(arrKeyValue,arrKeyIndex,object) 'call to drag drop function
        Case Else
            errStr1 = "Keyword check at Line No. -" &intRowCount
			errStr2 = "The keyword -'" & initial & "' is not supported.Please verify the keyword entered."
			err.raise vbObjectError
		End Select 'End of perform,storevalue, Checking	
End Function
'#################################################################################################

'#################################################################################################
'Function name 	: Func_Context_Win
'Description    : This function is used to set the full hierarchical Path for the object on which 
'				some action is to be performed.
'Parameters     : The Object details as the full hierarchical Path of the Object goes as parameter 
'				to the function and the current row number in the local Sheet.           
'Assumptions    : AUT is already up and running. 
'#################################################################################################
'The following function is for 'Context' keyword.
'#################################################################################################
Function  Func_Context_Win(arrObj,intRowCount)
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
			Select Case LCase(arrObj(0)) 'sets the heirarchical  path of the object
				Case "window" 'Parent is Window (e.g. error messages)
					Set curParent = Window(arrObj(1)) 'initial declaration of the parent object
				Case "dialog" 'Parent is dialog
					Set curParent = Dialog(arrObj(1)) 'initial declaration of the parent object
				 Case Else
						errStr1 = "Keyword Check  for 'Context' at Line no - " & intRowCount
						errStr2 = "Keyword - '" & arrObj(0) & "'  not supported.Please verify Keyword entered."
						err.raise vbObjectError
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
				Select Case LCase(arrChild(0)) 'to set child objects
					Case "dialog"
						Set curParent = curParent.Dialog(arrChild(1))
					Case "window"
						Set curParent = curParent.Window(arrChild(1))
					Case Else
						errStr1 = "Keyword Check  for 'Context' at Line no - " & intRowCount
						errStr2 = "Keyword - '" & arrChild(0) & "'  not supported.Please verify Keyword entered."
						err.raise vbObjectError
				End Select
			Next
            parChild = arrChild(0)'setting to check whether the parent is web or win
		End If
		Set parent = curParent 'Setting the current screen under which the object is present
		If  Environment.Value("icontext")=1 Then
		Call Func_CaptureScreenshot("test",intRowCount)   'Call to capture screenshot function
		wait 1
		End If
End Function
'#################################################################################################
'#################################################################################################
'Function name 	  : Func_ObjectSet_Win
'Description      : This function sets the parent and child objects.
'Parameters       : arrObjchk is an array of object names and intRowCount is the current Row number.
'Assumptions      : NA
'#################################################################################################
'The following function is called Internally 
'#################################################################################################
Function Func_ObjectSet_Win(arrObjchk,intRowCount)
   Call Func_DescriptiveObjectSet(arrObjchk,intRowCount)   
	arrObjchk(0) = Lcase(arrObjchk(0))
	If arrObjchk(0) <> "split" And arrObjchk(0) <> "random" And arrObjchk(0)<>"sqlvaluecapture" And arrObjchk(0)<>"sqlexecute" And arrObjchk(0)<>"sqlcheckpoint" And arrObjchk(0)<> "sqlmultiplecapture" And arrObjchk(0) <> "wait" Then
		Select Case LCase (arrObjchk(0)) 'Sets the parent and child objects
			Case "window"
				Set object = parent
			Case "dialog"
				Set object = parent
			Case "listbox"
				Set object = parent.WinList(arrObjchk(1))
			Case "spinner"
				Set object = parent.WinSpin(arrObjchk(1))
			Case "toolbar"
				Set object = parent.WinToolbar(arrObjchk(1))
			Case "button"
				Set object = parent.WinButton(arrObjchk(1))
			Case "treeview"
				Set object = parent.WinTreeView(arrObjchk(1))
			Case "listview"
				Set object = parent.WinListView(arrObjchk(1))
			Case "menu"
				Set object = parent.WinMenu(arrObjchk(1))
			Case "object"
				Set object = parent.WinObject(arrObjchk(1))
			Case "textbox"
				Set object = parent.WinEdit(arrObjchk(1))
		   Case "editor"
				Set object = parent.WinEditor(arrObjchk(1))
			Case "checkbox"
				Set object = parent.WinCheckbox(arrObjchk(1))
			Case "radiobutton"
				Set object = parent.WinRadiobutton(arrObjchk(1))
			Case "combobox"
				Set object = parent.WinComboBox(arrObjchk(1))
			Case "calendar"
				Set object = parent.WinCalendar(arrObjchk(1))
			Case "static"
				Set object = parent.Static(arrObjchk(1))
			Case "statusbar"
				Set object = parent.WinStatusBar(arrObjchk(1))
			Case "scrollbar"
				Set object = parent.WinScrollbar(arrObjchk(1))
			Case "tab"
				Set object = parent.WinTab(arrObjchk(1))
			Case "activex"
				Set object = parent.ActiveX(arrObjchk(1)) 
			Case Else
				errStr1 = "Keyword Check at Line no - " & intRowCount
				errStr2 = "Object  '"& arrObj(0)&"' is not supported.Please verify keyword entered."
				err.raise vbObjectError
			  Exit Function
		End Select
		Call Func_Wait_Win(arrObjchk,arrKeyValue,initial)
      End If 
End Function
'##############################################################################################
'##############################################################################################
'Function name 		: Func_Perform_Win
'Description        : If the user requires to perform a set of operations then the user can use this function
'Parameters       	: 1. Object on which the specified operation needs to be performed
'				  	  2. The operation that needs to be performed on the object                                 
'				  	  3. Additional parameters if required to identify the object where operation needs to be performed.                                
'Assumptions     	: NA
'##############################################################################################
'The following function is for Perform Keyword
'##############################################################################################
Function Func_Perform_Win(object,arrObj,arrKeyValue,arrKeyIndex,intRowCount)
    Select Case LCase(Trim(arrObj(0))) 'To perform  action on the object
		Case "window"	
			Select Case LCase(arrKeyValue(0)) 'to perform the operation on the windows object
				Case "close"	
					If object.GetROProperty("minimized") Then
						parent.Restore
					Else	
						parent.Activate
					End If
					parent.Close
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "activate"
					parent.Activate
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "maximize"
					If parent.GetROProperty ("enabled") Then'Check to see if a window is able to be maximized.
					   If parent.GetROProperty ("maximizable") Then
						 parent.Maximize
						 Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
         					   Else
					    Reporter.ReportEvent micFail, "Keyword Check  for 'Perform' at Line no - " & intRowCount, "Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
					   End If
					End If
				Case "minimize"
					If parent.GetROProperty ("enabled") Then 'Check to see if a window is able to be minimized.
					   If parent.GetROProperty ("minimizable") Then
						  parent.Minimize
						  Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					   End If
					End If
				Case "restore"
					parent.Restore
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "textclick"
					Call Func_SelectText(arrKeyIndex(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "dialog"
			Select Case LCase(arrKeyValue(0)) ' To perform the operation on the dialog objects
				Case "close"
					parent.Close
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "restore"	
					parent.Restore
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "activate"
					parent.Activate
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "textclick"
					Call Func_SelectText(arrKeyIndex(1))	
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "object"
			Select Case LCase(arrKeyValue(0)) 'Perform the operation on the object
				Case "type"
					object.Type arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "click"
					object.Click
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
				Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "button"
			object.Click
			Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0) &arrObj(1) 
		Case "toolbar"
			Select Case LCase(arrKeyValue(0)) 'To perform the operation on the tool bar
				Case "press"
					object.Press arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "spinner"
			Select Case LCase(arrKeyValue(0)) 'To perfrom the operation on spinner
				Case "next"
					object.Next
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "previous"
					object.Prev
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "set"
					object.Set CInt(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "click"
					object.Click
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "treeview"				
			Select Case LCase(arrKeyValue(0)) 'To perform operation on treeview
				Case "expand"
					object.Expand arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "expandall"			
					object.ExpandAll arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "collapse"
					object.Collapse arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "select"
					object.Select arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
                Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "listview"
			Select Case LCase(arrKeyIndex(0)) 'To perform operations on listview
				Case "select"
					object.Select arrKeyIndex(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "extendselect"
					object.ExtendSelect arrKeyIndex(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "deselect"
					object.Deselect arrKeyIndex(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "selectindex"
					object.Select CInt(arrKeyIndex(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "extendselectindex"
					object.ExtendSelect CInt(arrKeyIndex(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "deselectindex"
					object.Deselect CInt(arrKeyIndex(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "selectrange"
					object.SelectRange arrKeyIndex(1),arrKeyIndex(2)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "selectrangeindex"
					object.SelectRange CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "check"
					object.SetItemState arrKeyIndex(1),micChecked
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "uncheck"
					object.SetItemState arrKeyIndex(1),micUnChecked
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "itemclick"
					object.SetItemState arrKeyIndex(1),micClick
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "doubleclick"
					object.SetItemState arrKeyIndex(1),micDblClick
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "activate"
					object.Activate arrKeyIndex(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "click"
					object.Click
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyIndex(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "menu"
			object.Select arrKeyValue(0)
			Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		Case "combobox"
			Select Case LCase(arrKeyValue(0)) ' to perform the operation on combobox object
				Case "type"
					object.Type arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "select"	
					object.Select arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "selectindex"
					object.Select CInt(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "click"
					object.Click	
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "listbox"				
			Select Case LCase(arrKeyIndex(0))'to perform the operation on listbox
				Case "select"		
					object.Select arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "selectindex"
					object.Select CInt(arrKeyIndex(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "selectrange"
					If UBound(arrKeyIndex)>1 Then
						object.SelectRange arrKeyIndex(1),arrKeyIndex(2)
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
					Else	
						object.SelectRange arrKeyIndex(1)
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
					End If
				Case "selectrangeindex"
					If UBound(arrKeyIndex) > 1 Then
						object.SelectRange CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2))
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
					Else
						object.SelectRange CInt(arrKeyIndex(1))
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
					End If
				Case "deselect"
					object.DeSelect arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "deselectindex"
					object.DeselectSelect CInt(arrKeyIndex(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "extendselect"
					object.ExtendSelect arrKeyIndex(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "extendselectindex"	
					object.ExtendSelect CInt(arrKeyIndex(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action "& arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "click"
					object.Click
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)	
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyIndex(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "checkbox"
			Select Case LCase(arrKeyValue(0)) 'To perform the operation on checkbox object
				Case "check"
					object.Set "ON"
					object.WaitProperty "checked","ON",60000
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "uncheck"
					object.Set "OFF"
					object.WaitProperty "checked","OFF",60000
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "radiobutton"
			Select Case Trim(CStr(LCase(arrKeyValue(0)))) 'To perform the operation on radiobutton
				Case "set"
					object.Set
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "click"
					object.Click
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "tab"
			Select Case Trim(CStr(LCase(arrKeyValue(0)))) 'To perform the operation on tab
				Case "select"
					object.Select arrKeyValue(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "click"
					object.Click
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "selectindex"
					object.Select CInt(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "scrollbar"
			Select Case LCase(arrKeyValue(0)) 'To perform the operation on scrollbar
				Case "nextline"
					If UBound(arrKeyValue) > 0 Then
						object.NextLine CInt(arrKeyValue(1))
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					Else
						object.NextLine
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					End If			
				Case "nextpage"
					If UBound(arrKeyValue) > 0 Then
						object.NextPage CInt(arrKeyValue(1))
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					Else
						object.NextPage
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					End If
				Case "prevpage"
					If UBound(arrKeyValue) > 0 Then
						object.PrevPage CInt(arrKeyValue(1))
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					Else
						object.Prevpage
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					End If
				Case "prevline"	
					If UBound(arrKeyValue) > 0 Then
						object.PrevLine CInt(arrKeyValue(1))
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					Else
						object.PrevLine
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					End If
				Case "set"	
					object.Set CInt(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "mousemove"
					object.MouseMove Cint(arrKeyValue(1)),Cint(arrKeyValue(2))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0)&" performed successfully on "&arrObj(0)	
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select

   		Case "editor"
			Select Case Trim(CStr(LCase(arrKeyIndex(0)))) 'to perforn the operation on editor
				Case "type"
					object.Type arrKeyIndex(1)
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "setselection"
					object.SetSelection CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2)),CInt(arrKeyIndex(3)),CInt(arrKeyIndex(4))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "setcaretpos"
					object.SetCaretPos CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)
				Case "click"
					object.Click
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyIndex(0), "Action " &arrKeyIndex(0) &" performed successfully on "&arrObj(0)	
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyIndex(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "textbox"
			Dim arrStatus
			Dim propSplit1	
			Select Case Trim(CStr(LCase(arrKeyValue(0)))) 'to perform the operation on textbox
				Case "type"
					If (Instr(1,arrKeyValue(1),"d_") <> 0) Then
					propSplit1 = Split(arrKeyValue(1),"_",-1,1)
					arrStatus = Split(propSplit1(1),";",-1,1)
					VarName = arrStatus(0)
					Select Case LCase(arrStatus(0)) 'performing time and date operation
						Case "currenttime"
							Environment.Value(VarName) = FormatDateTime(Now(),4)
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
					  arrKeyValue(1) = Environment.Value(VarName)
				End If     
					object.Click
					object.type Trim(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "set"
					object.Set Trim(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "setsecure"
					object.SetSecure Trim(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "click"
					object.Click
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case "setselection"
					object.SetSelection CInt(arrKeyValue(1)),CInt(arrKeyValue(2))
				Case "doubleclick"
					object.DblClick 0,0
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Case Else
					Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Keyword - '" & arrKeyValue(0) & "'  not supported.Invalid '" & arrObj(0) &"' operation. Please verify keyword entered."
			End Select
		Case "calendar"
			Select Case Trim(CStr(LCase(arrKeyValue(0)))) 'to perform the operation for calendar
				Case "setdate"
					Select Case LCase(arrKeyValue(1))
						Case "now"
							object.SetDate Now
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" " &arrKeyValue(1) &" performed successfully on "&arrObj(0)
						Case "date"
							object.SetDate Date
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" " &arrKeyValue(1) &" performed successfully on "&arrObj(0)
						Case Else
							If IsDate(arrKeyValue(1)) Then
								object.SetDate CDate(arrKeyValue(1))
								Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0)  &" performed successfully on "&arrObj(0)
							Else
								Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Select Date. Invalid date provided."
							End If			
					End Select
				Case "settime"
					Select Case LCase(arrKeyValue(1)) ' to perform operation for system time
						Case "now"
							Dim curTime
							curTime = Time()
							object.SetTime curTime
							Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(1), "Action " &arrKeyValue(1)  &" performed successfully on "&arrObj(0)
						Case Else
							If TypeName(arrKeyValue(1)) = "Date" Then
								object.SetTime TimeValue(arrKeyValue(1))
								Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(1), "Action " &arrKeyValue(1)  &" performed successfully on "&arrObj(0)
							Else
								Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Select Date. Invalid date provided."
							End If
					End Select
				Case "click"	
					object.Click
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			End Select
		Case Else											
			If (LCase(Trim(arrObj(0)))= "sqlexecute") Or (LCase(Trim(arrObj(0)))= "sqlvaluecapture") Or (LCase(Trim(arrObj(0)))= "sqlcheckpoint") Or (LCase(Trim(arrObj(0)))= "sqlmultiplecapture") Then
				Environment("connectionString") = "DRIVER={Microsoft ODBC for Oracle};UID=" & Environment("dbUID") & ";PWD=" & Environment("dbPWD") & ";SERVER=" & Environment("dbServer") & "_" & Environment("dbHost") & ".world"	'Assign the Environment value to the Environment variable 'connectionString'
				Select Case LCase(Trim(arrObj(0))) ' to perform sql operations
					Case "sqlexecute"
						strSQL = arrObj(1)
						For inti = 0 to (Func_RegExpMatch ("##\w*##", arrObj(1),aPosition,aMatch) - 1)
							strReplace = "'" & Environment.Value(Replace(aMatch(inti),"##","",1,-1,1)) & "'"
							strSQL = Func_gfRegExpReplace(aMatch(inti), strSQL, strReplace)
						Next							
						Set dbConn = CreateObject("ADODB.Connection")	'Create the DB Object
						dbConn.Open Environment.Value("connectionString")							
						Set dbRs = dbConn.Execute(strSQL)	'Execute the query
						dbConn.Close   ' Close the database connection
						Set dbConn = Nothing	
						Reporter.ReportEvent micDone, " Sql Operation", "Query executed successfully" 						
					Case "sqlvaluecapture"
						Environment.Value(arrKeyValue(0)) = Func_gfQuery(arrObj(1))
						Reporter.ReportEvent micDone, " Sql value capture", "Sql value captured successfully" 
					Case "sqlmultiplecapture"
						DataTable.GetSheet("Action1").SetCurrentRow(1)
						DbTable(arrKeyIndex(0)).SetTOProperty "connectionstring", Environment.Value("connectionString")	'Set the TO property of connection string.
						strSQL = arrAction(1)
						For inti = 0 to (Func_RegExpMatch ("##\w##", arrAction(1),aPosition,aMatch) - 1)
							strReplace = "'" & Environment.Value(Replace(aMatch(inti),"##","",1,-1,1)) & "'"
							strSQL = Func_gfRegExpReplace(aMatch(inti), strSQL, strReplace)
						Next
						DbTable(arrKeyIndex(0)).SetTOProperty "source", strSQL	'Change the DB OBjects source(SQL) statement.
						DbTable(arrKeyIndex(0)).Output CheckPoint(arrKeyIndex(0)) 'Execute the DB output checkpoint.
						Reporter.ReportEvent micDone, " sql multiple capture", "sql multiple captured successfully" 
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
				Reporter.ReportEvent micFail,"Keyword Check  for 'Perform' at Line no - " & intRowCount,"Object  '"&arrObj(0)&"' is not supported.Please verify keyword entered."	 
				If htmlreport = "1"  Then
					Call Update_log(MAIN_FOLDER, g_sFileName, "fail")  ' calling function update log to create an execution log in HTML file
				End If
		End If	
	End Select 
		If   Environment.Value ("iperform") =1 Then
		Call Func_CaptureScreenshot("test",intRowCount)  ' Call to capture screenshot function
		wait 1
	End If
End Function
'#################################################################################################

'#################################################################################################
'Function name 		: Func_Store_Win
'Description    	: If the user requires to store any property of a particular object into a variable 
'					  then this function can be used.
'Parameters     	: The Object details as the full hierarchical Path of the Object goes as parameter 
'					  to the function.            
'Assumptions    	: None 
'#################################################################################################
'The following function is for StoreValue keyword.
'#################################################################################################
Function Func_Store_Win(object)
   'Splits the data present in the Fourth Column into the property to be stored and the variable in which to store. 
   propSplit = Split(Datatable.Value(4,dtLocalSheet),":",2,1)
	propName = propSplit(0)
	VarName = propSplit(1)
	Select Case LCase(propName) 'To store the value of a object into a variable
		Case "itemcount"
			Environment.Value(VarName) = object.GetROProperty("items count")
		Case "columncount"
			Environment.Value(VarName) = object.ColumnCount
		Case "disabled"
			Environment.Value(VarName) = Not(CBOOL(object.GetROProperty("enabled")))	
		Case "rowcount"
			Environment.Value(VarName) = object.RowCount
		Case "exist"
			 Environment.Value(VarName) = object.Exist
		Case "visibletext"	 
			Environment.Value(VarName) = object.GetVisibleText	
		Case Else
			Environment.Value(VarName) = object.GetROProperty(propName)	
	End Select

				  If cint(Introwcount)<=cint(Environment("intEndRow")) Then
						If Cint(Environment("intStartRow"))<=Cint(Introwcount) Then
							Call DebugGetEnv()        'to call debug function for execution log status in HTML file
						End If  
				 End If

End Function
'#################################################################################################
'#################################################################################################
'Function name 	: Func_Check_Win
'Description    : This function is used for all the checking operations to be performed on the AUT.
'Parameters     : The Object details on which check has to be performed along with details of the 
'				  fourth Column of the current row in Local sheet and the current row number in the 
'				  Local Sheet.           
'Assumptions    : NA
'#################################################################################################
'The following function is for 'Check' keyword.
'#################################################################################################
Function Func_Check_Win(object,arrObjchk,arrKeyValue,arrKeyIndex,intRowCount)
	Dim strChkParameter	'variable is used to store the mode of checking
	Dim ActualValue  'stores the actual property value of the object at run time
	Dim ExpectedValue   'stores the value of the property of the object defined in datatable
	Dim strStatus  'stores the status of the execution
	Dim iStatus   ' stores the status of the object as 'pass' or 'fail'
	Dim reportStep  'stores the step to be performed on the object
	Dim reportStepPass 'stores the pass result of the object 
	Dim reportStepFail 'stores the fail result of the object
	Dim result  'Stores the return value (‘true’ or ‘false’) 
	Dim checkval   'initialization variable used for func_checkvaluereturn
	checkval = 2
    strChkParameter = "exactchk"
	iStatus = "iDone"
	If UBound(arrObjchk)> 1 Then
		strChkParameter= arrObjchk(2)
	End If
	curObjClassName = arrObjchk(1)
	arrKeyValue(1) = GetValue(arrKeyValue(1))
    	
		Select Case LCase (arrKeyValue(0)) 'to perform check operation of the object

			Case "enabled"
				If LCase(arrKeyValue(1)) = "true" Or LCase(arrKeyValue(1)) = "false" Then
				ActualValue = CBool(object.GetROProperty(LCase(arrKeyValue(0))))
				ExpectedValue = CBool(arrKeyValue(1))
					If LCase(ExpectedValue) = "false" Then
					strStatus = " not "
					Else
					strStatus = ""
					End If
					reportStep = "Verify that '" & curObjClassName & "'  "& arrObjchk(0) &" is "&  strStatus &" " &arrKeyValue(0) & "."
                    If ExpectedValue = ActualValue Then
					reportStepPass = "The '"& curObjClassName &"'  "& arrObjchk(0) &" is " & strStatus & " " & arrKeyValue(0) & ", which is as expected."
					iStatus = "iPass"
					keyword = 0
					checkval=1
						
				Else
					If LCase(ExpectedValue) = "false" Then
						reportStepFail = "The '"& curObjClassName &"'  "& arrObjchk(0) &" is " & arrKeyValue(0) & ", which is not as expected."
					Else
						reportStepFail = "The '"& curObjClassName &"'  "& arrObjchk(0) &" is  not " & arrKeyValue(0) & ", which is not as expected."	
					End If
					iStatus = "iFail"
					keyword = 1
					checkval=0

				End If   		 
		End If	
		
			Case "focused"
				If LCase(arrKeyValue(1)) = "true" Or LCase(arrKeyValue(1)) = "false" Then
				ActualValue = CBool(object.GetROProperty(LCase(arrKeyValue(0))))
				ExpectedValue = CBool(arrKeyValue(1))
					If LCase(ExpectedValue) = "false" Then
					strStatus = " not "
					Else
					strStatus = ""
					End If
			reportStep = "Verify that '" & curObjClassName & "'  "& arrObjchk(0) &" is "&  strStatus &" " &arrKeyValue(0) & "."			
               If ExpectedValue = ActualValue Then
					reportStepPass = "The '"& curObjClassName &"'  "& arrObjchk(0) &" is " & strStatus & " " & arrKeyValue(0) & ", which is as expected."
					iStatus = "iPass"
					keyword = 0
					checkval=1
				Else
					If LCase(ExpectedValue) = "false" Then
						reportStepFail = "The '"& curObjClassName &"'  "& arrObjchk(0) &" is " & arrKeyValue(0) & ", which is not as expected."
					Else
						reportStepFail = "The '"& curObjClassName &"'  "& arrObjchk(0) &" is  not " & arrKeyValue(0) & ", which is not as expected."	
					End If
					iStatus = "iFail"
					keyword = 1
					checkval=0
				End If
		End If	
		
			Case "visible"
				If LCase(arrKeyValue(1)) = "true" Or LCase(arrKeyValue(1)) = "false" Then
				ActualValue = CBool(object.GetROProperty(LCase(arrKeyValue(0))))
				ExpectedValue = CBool(arrKeyValue(1))
					If LCase(ExpectedValue) = "false" Then
					strStatus = " not "
					Else
					strStatus = ""
					End If
			reportStep = "Verify that '" & curObjClassName & "'  "& arrObjchk(0) &" is "&  strStatus &" " &arrKeyValue(0) & "."			
               If ExpectedValue = ActualValue Then
					reportStepPass = "The '"& curObjClassName &"'  "& arrObjchk(0) &" is " & strStatus & " " & arrKeyValue(0) & ", which is as expected."
					iStatus = "iPass"
					keyword = 0
					checkval=1
				Else
					If LCase(ExpectedValue) = "false" Then
						reportStepFail = "The '"& curObjClassName &"'  "& arrObjchk(0) &" is " & arrKeyValue(0) & ", which is not as expected."
					Else
						reportStepFail = "The '"& curObjClassName &"'  "& arrObjchk(0) &" is  not " & arrKeyValue(0) & ", which is not as expected."	
					End If
					iStatus = "iFail"
					keyword = 1
					checkval=0
				End If		   		 
		End If	


			Case "itemcount"
				ExpectedValue = CInt(arrKeyValue(1))
				ActualValue  = CInt(object.GetROProperty("items count"))
				reportStep = "Verify the number of items in  '"& curObjClassName &"'  "& arrObjchk(0) &" is "& arrKeyValue(1) & "."
				If Lcase(ActualValue)=Lcase(ExpectedValue) Then
				reportStepPass = "The number of items in '"& curObjClassName &"'  "& arrObjchk(0) &" is "& ActualValue & ", which is as expected."
				istatus="iPass"
				keyword=0
				checkval=1
				else
				reportStepFail =  "The number of items in '"& curObjClassName &"'  "& arrObjchk(0) &" is "& ActualValue & ", which is not as expected of '"&arrKeyValue(1)&"' ."
				istatus="iFail"
				keyword=1
				checkval=0
				End If

			Case "columncount"
				ExpectedValue = CInt(arrKeyValue(1))
				ActualValue  = CInt(object.ColumnCount)
				If Lcase(ActualValue)=Lcase(ExpectedValue) Then
				reportStepPass = "The number of columns in '"& curObjClassName &"'  "& arrObjchk(0) &" is "& ActualValue & ", which is as expected."
				istatus="iPass"
				keyword=0
				checkval=1
				Else
				reportStepFail =  "The number of columns in '"& curObjClassName &"'  "& arrObjchk(0) &" is "& ActualValue & ", which is not as expected of '"&arrKeyValue(1)&"' ."
				istatus="iFail"
				keyword=1
				checkval=0
				End If

			Case "rowcount"
				ExpectedValue = CInt(arrKeyValue(1))
				ActualValue  = CInt(object.RowCount)
				reportStep = "Verify the number of rows in  '"& curObjClassName &"'  "& arrObjchk(0) &" is "& arrKeyValue(0) & "."
				If Lcase(ActualValue)=Lcase(ExpectedValue) Then
				reportStepPass = "The number of rows in '"& curObjClassName &"'  "& arrObjchk(0) &" is "& ActualValue & ", which is as expected."
				istatus="iPass"
				keyword=0
				checkval=1
				Else
				reportStepFail =  "The number of rows in '"& curObjClassName &"'  "& arrObjchk(0) &" is "& ActualValue & ", which is not as expected of '"&arrKeyValue(1)&"' ."
				istatus="iFail"
				keyword=1
				checkval=0
				End If

			Case "text"
               	ExpectedValue = Trim(arrKeyValue(1))
				ActualValue  = Trim(object.GetROProperty("text"))
				reportStep = "Verify that  "& arrKeyValue(0) & " displayed in '"& curObjClassName &"'  "& arrObjchk(0)&" is "& arrKeyValue(1) & "."
				If Lcase(ExpectedValue) = Lcase(ActualValue) Then
                reportStepPass = "The text displayed in '"& curObjClassName &"'  "& arrObjchk(0) &" is "& ActualValue & ", which is as expected."
				istatus="iPass"
				keyword = 0
				checkval =1
				else
                reportStepFail =  "The text displayed in '"& curObjClassName &"'  "& arrObjchk(0) &" is "& ActualValue & ", which is not as expected of '"&arrKeyValue(1)&"' ."		
				istatus="iFail"
				 keyword = 1
				checkval = 0
				End If

			Case "selection"
				ExpectedValue = Trim(arrKeyValue(1))
				ActualValue  = Trim(object.GetROProperty("selection"))
				reportStep = "Verify that  "& arrKeyValue(1) & " is selected in '"& curObjClassName &"'  "& arrObjchk(0) & "."
				If Lcase(ActualValue)=Lcase(ExpectedValue) Then
				reportStepPass = "The selected item in '"& curObjClassName &"'  "& arrObjchk(0) &" is " & ActualValue & ", which is as expected."
				istatus="iPass"
				keyword = 0
				checkval = 1 
				Else
				reportStepFail =  "The selected item in '"& curObjClassName &"'  "& arrObjchk(0) &" is "& ActualValue & ", which is not as expected of '"&arrKeyValue(1)&"' ."
				istatus="iFail"
				keyword = 1
				checkval=0
				End If

			Case "exist"
				If LCase(arrKeyValue(1)) = "true" Or LCase(arrKeyValue(1)) = "false" Then
					ActualValue = CBool(object.Exist)
					ExpectedValue = CBool(arrKeyValue(1))
					If object.Exist  Then						
						If Lcase(arrKeyIndex(1)) = "true"  Then
							reportstep="To verify " &arrObjchk(1) & "of " &arrObjchk (0) &" is "&arrKeyValue(0)
							reportStepPass = "The " & arrObjchk(1) & " of  '"& arrObjchk(0) &"'   exists  which is as, expected."
							iStatus="iPass"
							keyword = 0
							checkval=1
						ElseIf Lcase(arrKeyIndex(1)) = "false" Then
							reportstep="To verify " &arrObjchk(1) & "of " &arrObjchk (0) &" is not "&arrKeyValue(0)
							reportStepFail = "The " & arrObjchk(1) & " of  '"& arrObjchk(0) &"'   exists, which is not as expected."
							iStatus="iFail"
							keyword = 1
							checkval=0
                        	End If
					Else
					End If
				End If	

			Case "checked"
				If LCase(arrKeyValue(1)) = "on" Or LCase(arrKeyValue(1)) = "off" Then
					ActualValue = object.GetROProperty("checked")
					ExpectedValue = UCase(arrKeyValue(1))
					If LCase(ExpectedValue) = "off" Then
						strStatus = " unchecked "
					Else
						strStatus = "checked"
					End If
					reportStep =  "Verify that '" & curObjClassName & "'  "& arrObjchk(0) &" is "&  strStatus & "."
					If Lcase(ActualValue)=Lcase(ExpectedValue) Then
					reportStepPass = "The " & arrObjchk(0) & " '"& arrObjchk(1) &"'  is  "&  strStatus & ", which is as expected."
					istatus="iPass"
					keyword=0
					checkval=1
					Else
					reportStepFail =  "The " & arrObjchk(0) & " '"& arrObjchk(1) &"' is not "&  strStatus & ", which is not as expected."
					istatus="iFail"
					keyword=1
					checkval=0
					End If
				End If

			Case "tabexist"
				Dim itemFound
				Dim actualCount
				itemFound = 2
				actualCount = object.GetROProperty("items count")
				For inti = 0 to actualCount-1
					ActualValue = object.GetItem(inti)
					If Lcase(strChkParameter) = "regexpchk" Then
						If Instr(1,ActualValue,arrKeyValue(1)) <> 0 Then
							itemFound = 0
							Exit For
						End If
					Else
						If Trim(ActualValue) = Trim(arrKeyValue(1)) Then
							itemFound =0
							Exit For
						End If
					End If					
				Next
				reportStep =  "Verify that Tab item '" & arrKeyValue(1) & "' is present in the Tab '" & arrObjchk(1) &"'."
				If itemFound <> 0 Then   
					reportStepFail = "The Tab item  '"&arrKeyValue(1)&"' does not exist in the Tab '" & arrObjchk(1) & "', which is not as expected."
					iStatus = "iFail"
					keyword = 1
					checkval=0
				Else
					reportStepPass = "The Tab item  '"&arrKeyValue(1)&"' exists in the Tab '" & arrObjchk(1) & "and at position '" & inti & "', which is as expected."
					iStatus = "iPass"
					keyword = 0
					checkval = 1
				End If

			Case "itemexist"
                 itemFound = 2
				actualCount = CInt(object.GetItemsCount)
				 For inti = 0 to actualCount-1
					ActualValue = object.GetItem(inti)
					If Lcase(strChkParameter) = "regexpchk" Then
						If Instr(1,ActualValue,arrKeyValue(1)) <> 0 Then
							itemFound = 0
							Exit For
						End If
					Else
						If Trim(ActualValue) = Trim(arrKeyValue(1)) Then
							itemFound =0
							Exit For
						End If
					End If					
				Next
				reportStep =  "Verify that item '" & arrKeyValue(1) & "' is present in the '" & arrObjchk(0) &"."
				If itemFound <> 0 Then
					reportStepFail = "The item  '"&arrKeyValue(1)&"' does not exist in the " & arrObjchk(0) & ", which is not as expected."
					iStatus = "iFail"
					keyword = 1
					checkval=0
				Else
					reportStepPass = "The item  '"&arrKeyValue(1)&"' exists in the " & arrObjchk(0) & " and at position '" & inti & "', which is as expected."
					iStatus = "iPass"
					keyword = 0
					checkval=1
				End If
					  
			Case "windowtext"
				Dim strText  'stores the text to be in search for
				Dim intlen  'initialization variable
				Dim inttex   'initialization variable
				Dim introw  'stores the width of the object
				Dim intbre  'stores the height of the object
				 intlen = 2
				 inttex = 2
				 introw = CInt(object.GetROProperty("width")) - 10
				 intbre = CInt(object.GetROProperty("height")) - 10
				strText = object.GetVisibleText(intlen,inttex,introw,intbre)
				'Replace the Carriage Return Character
				strText = Replace(strText,Chr(13),"",1,-1,1)
				'Replace the New Line Character
				strText = Replace(strText,Chr(10),"",1,-1,1)
				'Remove any spaces.
				strText = Replace(strText," ","",1,-1,1)
				'Need to handle an optional True/False parameter.
				'Assume True if the optional parameter is not supplied.
				ActualValue = True
				If UBound(arrKeyIndex) >1 Then
					ExpectedValue = CBool(arrKeyIndex(2))
				End If
				ActualValue= Func_gfRegExpTest(arrKeyIndex(1), strText)
				reportStep= "Checkpoint - WindowText"
				If  LCase(ActualValue) =LCase(ExpectedValue) Then
				reportStepPass=  "Baseline:  " & arrKeyIndex(1) & vbCrLf & "Exist:  " & ActualValue & vbCrLf & "Actual:  " & strText & ", which is  as expected."
				istatus="iPass"
				keyword=0
				checkval=1
				Else If  LCase(ExpectedValue)="false" Then
						reportStepFail = "Baseline:  " & arrKeyIndex(1) & vbCrLf & " Exist:  " & ActualValue & vbCrLf & "Actual:  " & strText & ", which is not as expected."
					Else
						reportStepFail = "Baseline:  " & arrKeyIndex(1) & vbCrLf & "Not Exist :  " & ExpectedValue & vbCrLf & "Actual:  " & strText & ", which is not as expected."	
					End If
					iStatus = "iFail"
					keyword = 1
					checkval=0
				End If

				Case "tablesearch"
									searchtablereturn= Func_tablesearch(object,arrKeyValue(1))
									reportStep=	"Verify that a row value matches for search criteria '"&arrKeyValue(1)
								   arrsearchtablereturn=Split(searchtablereturn,":")
									If Lcase(arrsearchtablereturn(1)) =Lcase( arrsearchtablereturn(2))Then
											reportStepPass="The row value corresponding for the search criteria " &arrKeyValue(1)& "  is exists, which is as expected."
											Reporter.ReportEvent micPass,reportStep,reportStepPass
											  keyword=0
											 checkval=1
									Elseif 	Lcase(arrsearchtablereturn(1)) <> Lcase(arrsearchtablereturn(2)) Then
											reportStepFail =  "The corresponding row values doesn't suit the search criteria '"&arrKeyValue(1) &",  Which is  not as expected."											
                							Reporter.ReportEvent  micFail,reportStep,reportStepFail
											keyword=1
											checkval=0
									End If
      
				Case Else 'for checking the properties which is not listed in the above select case statement
						actualValue = CStr(object.GetROProperty(arrKeyIndex(0))) 
						expectedValue = (arrKeyIndex(1)) 
						reportStep= "To Verify the '" &arrKeyIndex(0) &"' property of " &arrObjchk(0) &" "&curObjClassName
						If Lcase( actualValue)=Lcase(expectedValue)Then
							reportStepPass= "The '" &arrKeyIndex(0) &"' property of "&arrObjchk(0) &" " &curObjClassName  &"' is "  &arrKeyIndex(1) &", which is as expected."
							iStatus = "iPass"
							keyword=0
							checkval=1
						Elseif Lcase(actualValue)="" Then 'checking for unavailable properties of an object
							reportStep =  "Verify the "&arrKeyIndex(0)&" property of '" & arrAction(1) & "'  "& arrAction(0) &"." 
							reportStepFail = "Property does not exists. Please verify property entered."
							Reporter.ReportEvent micFail,reportStep,reportStepFail
							Call Func_CaptureScreenshot("checkpoint",intRowCount)	   ' Call to capture screenshot function
						Exit Function
				Else
						reportStep =  "Verify the "&arrKeyIndex(0)&" property of '" & arrAction(1) & "'  "& arrAction(0) &"." 
						reportStepFail= "The '" &arrKeyIndex(0) &"' property of "&arrObjchk(0) &" " &curObjClassName  &"' is not "  &actualValue &", which is not as expected."
						iStatus = "iFail"
						Reporter.ReportEvent micFail,reportStep,reportStepFail
						keyword=1
						checkval=0
						Call Func_CaptureScreenshot("checkpoint",intRowCount)	' Call to capture screenshot function
						End if
				End Select   

		If iStatus = "iDone" Then
			If  (LCase(Trim(strChkParameter))= "exactchk") Then
				ActualValue = Cstr(Trim(ActualValue))
				ExpectedValue = Cstr(Trim(ExpectedValue))
				result= Eval("actualValue = expectedValue")
			Else
				ActualValue = Cstr(Trim(ActualValue))
				ExpectedValue = Cstr(Trim(ExpectedValue))
				result= Func_gfRegExpTest(ExpectedValue, ActualValue)
			End If
			If result Then
				Reporter.ReportEvent micPass,reportStep,reportStepPass
				iStatus = "iPass"
				keyword = 0
			Else
				Reporter.ReportEvent micFail,reportStep,reportStepFail
				iStatus = "iFail"
				keyword = 1
				Call Func_CaptureScreenshot("checkpoint",intRowCount)	 ' Call to capture screenshot function
			End If
		ElseIf  iStatus = "iPass" Then
				Reporter.ReportEvent micPass,reportStep,reportStepPass  
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	' calling function update log to create an execution log in HTML file
					End If	
		ElseIf  iStatus = "iFail" Then
				Reporter.ReportEvent micFail,reportStep,reportStepFail
					If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	   ' calling function update log to create an execution log in HTML file
					End If 
				Call Func_CaptureScreenshot("checkpoint",intRowCount)	 ' Call to capture screenshot function
		End If

		 If Datatable.Value(5,dtLocalSheet) <> empty Then
				Call Func_CheckValueReturn(intRowCount,checkval)
			 If cint(Introwcount)<=cint(Environment("intEndRow")) Then
			    	If Cint(Environment("intStartRow"))<=Cint(Introwcount) Then
				   Call DebugGetEnv()         'to call debug function for execution log status in HTML file
				   End If	
	     End If
			End If			

		If   Environment.Value("icheck")=1  Then
		Call Func_CaptureScreenshot("test",intRowCount)   ' Call to capture screenshot function
		wait 1
	End If

End Function
'#####################################################################################################
'Function name 	: Func_DragDrop_Win
'Description        : Drags and Drop the object from the original position to the specified position
'Parameters       	: 		arrKeyValue:value from 4th column of datatable 
'										Object	 :: This is the object on which the specified operation needs to be performed.  .
'Assumptions     	: NA
'#####################################################################################################
Function Func_DragDrop_Win(arrKeyValue,object)
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
'########################################################################################################
'####################################################################################################################
'Function name 	    : Func_Wait_Win
'Description        : This function is used for synchronization  with the application
'Parameters       	: The 'Object type' and the 'action being performed is passed as parameters. 
'Assumptions     	: None
'####################################################################################################################
'The following function is used internally.
'####################################################################################################################
Function Func_Wait_Win(arrObjchk,arrKeyValue,initial)
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
'##########################################################################################################