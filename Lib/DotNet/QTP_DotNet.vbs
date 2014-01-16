'###############################  KEYWORD DRIVEN FRAMEWORK   #####################################
'Project Name		: DotNet Framework
'Author		       	: Open2Test
'Version	    	: V 2.0
'Date of Creation	: 31-May-2012
'######################################  Driver Function  ################################################

'#################################################################################################
' Function Name 	: Keyword_DotNet
'Description        : This is the main function which interprets the specific keywords and performs the 
'					  desired actions. All the specific keywords used in the datatable are processed in this function.
'Parameters       	: Keyword in the 2nd column of the data table
'Assumptions     	: The Automation Script is present in the Local Sheet of QTP.
'#################################################################################################
'The following function is called internally
'#################################################################################################
Function Keyword_DotNet(initial)
   	 If htmlreport = "1" Then
		Call Update_Log(MAIN_FOLDER, g_sFileName,"executed")' calling function update log to create an execution log in HTML file
	 End If	
   On error resume next
	If initial = "context" Then
	   Call Func_Context_DotNet(arrAction,intRowCount)
	   Exit Function
	End If
	Set object = Nothing
	Call Func_ObjectSet_DotNet(arrAction,intRowCount)
	Select Case LCase(initial)' statement to do keyword operation
		Case "perform"
			'start perform
			Call Func_Perform_DotNet(object,arrAction,arrKeyValue,arrKeyIndex,intRowCount)
        Case "storevalue"
			Call Func_Store_DotNet(object,arrAction)
            'Checking
		Case "check"
			Call Func_Check_DotNet(object,arrAction,arrKeyValue,arrKeyIndex,intRowCount)
        Case Else
            errStr1 = "Keyword check at Line No. -" &intRowCount
			errStr2 = "The keyword -'" & initial & "' is not supported.Please verify the keyword entered."
			err.raise vbObjectError
		End Select 'End of perform,storevalue, Checking	
End Function
'#################################################################################################
'Function name 	    : Func_Context_DotNet
'Description    	: This function is used to set the full hierarchical Path for the object on which 
'					  some action is to be performed.
'Parameters     	: The Object details as the full hierarchical Path of the Object goes as parameter 
'		   			  to the function and the current row number in the Global Sheet.           
'Assumptions    	: AUT is already up and running. 
'#################################################################################################
'The following function is for 'Context' keyword.
'#################################################################################################
Function Func_Context_DotNet(arrObj,intRowCount)
    Dim arrChildCell	'stores the elements separated  by the delimiter '::'
    Dim contextData		'Stores the value present in the fourth Column
	Dim arrChild		'Stores the child objects of the main window
     inti = 0
	 Call Func_DescriptiveObjectSet(arrObj,intRowCount)
     	Select Case LCase(arrObj(0)) 'setting parent objects
		Case "window"
			Set curParent = SwfWindow(arrObj(1))
		Case "dialog"
			Set curParent = Dialog(arrObj(1))
		Case "browser"
			Set curParent = Browser(arrObj(1))
		Case "popupwindow"
			Set curParent = Window(arrObj(1))
		Case "vbwindow"
			Set curParent = vbWindow(arrObj(1))		
	End Select
	If (CStr(Trim(DataTable.Value(4, dtLocalSheet))) <> "") Then
		contextData = CStr(Trim(DataTable.Value(4, dtLocalSheet)))
		arrChildCell = Split(contextData, "::", -1, 1)
		For intj = 0 To UBound(arrChildCell)
			arrChild = Split(arrChildCell(intj), ";", 2, 1) 'checks the child object type
			If UBound(arrChildDesc) > 0 Then
			Call Func_DescriptiveObjectSet(arrChild,intRowCount)
			End If
			Select Case LCase(arrChild(0))  'setting child objects
				Case "dialog"
					Set curParent = curParent.Dialog(arrChild(1))
				Case "table"
					Set curParent = curParent.SwfTable(arrChild(1))
				Case "window"
					Set curParent = curParent.SwfWindow(arrChild(1))
				Case "popupwindow"
					Set curParent = objParPage.Window(arrChild(1))
				Case "vbwindow"
					Set curParent = objParPage.VbWindow(arrChild(1))
				Case Else
					 Reporter.ReportEvent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrChild(0) & "'  not supported.Please verify Keyword entered."
			End Select
		Next
		newContext = 1
		parChild = arrChild(0)'setting to check whether the parent is web or win
	End If
	Set parent = curParent 'Setting the current screen under which the object is present

		If  Environment.Value("icontext")=1 Then
			Call Func_CaptureScreenshot("test",intRowCount)    'call to screencapture function to take screenshot
			wait 1
		End If
End Function

'#################################################################################################
'#################################################################################################
'Function name 	  : Func_ObjectSet_DotNet
'Description      : This function sets the parent and child objects.
'Parameters       : arrObj is an array of object names and intRowCount is the current Row number.
'Assumptions      : NA
'#################################################################################################
'The following function is called Internally 
'#################################################################################################
Function Func_ObjectSet_DotNet(arrObj, intRowCount)
   Dim ObjectVal 		'Stores the Table object Class name
   Call Func_DescriptiveObjectSet(arrObj,intRowCount) 
   'If condition for object setting for objects other than 'Window', 'Dialog' and 'Browser'
	If  arrObj(0) <> "split" And arrObj(0) <> "random" And arrObj(0)<>"sqlvaluecapture" And arrObj(0)<>"sqlexecute" And arrObj(0)<>"sqlcheckpoint" And arrObj(0)<>"sqlmultiplecapture" Then 
		Select Case LCase(arrObj(0)) 'setting the parent objects
			Case "frame"
				Set object = parent.frame(arrObj(1)) 'initial declaration of the object
			Case "window"
				Set object = parent
			Case "browser"
				Set object = parent	
			Case "popupwindow"
				Set object = parent	
			Case "dialog"
				Set object = parent
			Case "listbox"
				Set object = parent.SwfList(arrObj(1))
			Case "spinner"
				Set object = parent.SwfSpin(arrObj(1))
			Case "toolbar"
				Set object = parent.SwfToolbar(arrObj(1))
			Case "wintoolbar"
	            Set object = parent.WinToolbar(arrObj(1))
			Case "button"
				Set object = parent.SwfButton(arrObj(1))
			Case "winbutton"
				Set object = parent.WinButton(arrObj(1))
			Case "vbbutton"
		         Set object = parent.VbButton(arrObj(1))
			Case "treeview"
				Set object = parent.SwfTreeView(arrObj(1))
			Case "label"
				Set object = parent.SwfLabel(arrObj(1))
			Case "listview"
				Set object = parent.SwfListView(arrObj(1))
			Case "menu"
				Set object = parent.WinMenu(arrObj(1))
			Case "object"
				Set object = parent.SwfObject(arrObj(1))
			Case "textbox"
				Set object = parent.SwfEdit(arrObj(1))
			Case "wintextbox"
				Set object = parent.WinEdit(arrObj(1))
			Case "editor"
				Set object = parent.SwfEditor(arrObj(1))
			Case "checkbox"
				Set object = parent.SwfCheckbox(arrObj(1))
			Case "radiobutton"
				Set object = parent.SwfRadiobutton(arrObj(1))
			Case "combobox"
				Set object = parent.SwfComboBox(arrObj(1))
			Case "static"
				Set object = parent.Static(arrObj(1))
			Case "statusbar"
				Set object = parent.SwfStatusBar(arrObj(1))
			Case "calendar"
				Set object = parent.SwfCalendar(arrObj(1))
			Case "scrollbar"
				Set object = parent.SwfScrollbar(arrObj(1))
			Case "tab"
				Set object = parent.SwfTab(arrObj(1))
			Case "table"
				Set object = parent.SwfTable(arrObj(1))
			Case "tabstrip"
				Set object = parent.WbfTabStrip(arrObj(1))
			Case "ultragrid"
				Set object = parent.WbfUltraGrid(arrObj(1))
			Case "webgrid"
				Set object = parent.WbfGrid(arrObj(1))
			Case Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrObj(0) & "'  not supported.Please verify Keyword entered."
		End Select 'End of object settings
		Call Func_Wait_DotNet(arrObj,arrKeyValue,initial)
	End If 'End of the setting object if condition
End Function
'##############################################################################################
'##############################################################################################
'Function name 		: Func_Perform_Dotnet
'Description     	: If User requires to perform a set of operations then the user can use this function
'Parameters       	: 1. Object on which the specified operation needs to be performed
'					  2. The operation that needs to be performed on the object                                 
'					  3. Additional parameters if required to identify the object where operation needs to be performed.                                
'Assumptions     	: NA
'##############################################################################################
'The following function is for Perform Keyword
'##############################################################################################
Function Func_Perform_DotNet(object, arrObj, arrKeyValue, arrKeyIndex, intRowCount)
	Select Case LCase(Trim(arrKeyValue(0)))' Selecting the specific action to be performed
		Case "tablesearch"
				Call Func_tablesearch(object,arrTableindex(1))
		Case "rownum"
		   	    strParam = CStr(Trim(DataTable.Value(5,dtLocalSheet)))
			Call func_getRowNum_Dotnet(object,arrKeyValue(1),strParam)
		Case "activate"
			curParent.Activate
			Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
        Case "activateitem" ' used for activating item in Listview
            	If arrObj(0) = "listview" Or arrObj(0) = "listbox"  Then
				object.Activate arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If			
        Case "activatecell" ' used for activating a cell in table
			If arrObj(0) = "table"   Then
				object.ActivateCell CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2))
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "close"
			If arrObj(0) = "window" Or  arrObj(0) = "dialog"   Then
					If curParent.GetROProperty("minimized") Then
						curParent.Restore
					Else
						curParent.Activate
					End If
						curParent.Close
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If								
		Case "click"			
			object.Click
			Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		Case "clickitem" ' used for activating item in Listview
			If arrObj(0) = "listview"   Then
				object.SetItemState arrKeyValue(1), micClick	
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0) 
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If	   	
		Case "collapse"
			If arrObj(0) = "treeview"   Then
				object.Collapse arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If	 			
		Case "check"
			If arrObj(0) = "treeview" Or arrObj(0) = "listview" Then
				object.SetItemState arrKeyValue(1), micChecked
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			ElseIf arrObj(0) = "checkbox" Then
				object.Set "ON"				 
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If		   
		Case "deselect"
			If arrObj(0) = "listbox" Or arrObj(0) = "listview"  Then
				object.Deselect arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "deselectindex"
			If arrObj(0) = "listbox" Or arrObj(0) = "listview"  Then
			object.Deselect CInt(arrKeyValue(1))
			Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If			
		Case "doubleclick"
			If arrObj(0) = "listview"  Then
            		object.SetItemState arrKeyValue(1), micDblClick
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			ElseIf arrObj(0) = "textbox" Then
			  object.DblClick 0,0
			  Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If	
        	Case "expand"
			If arrObj(0) = "treeview"   Then
				object.Expand arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If			
		Case "expandall"
			If arrObj(0) = "treeview"   Then
				object.ExpandAll arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			 Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "extendselect"
			If arrObj(0) = "listbox" Or arrObj(0) = "listview"  Then
				object.ExtendSelect arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			 Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "extendselectindex"
			If arrObj(0) = "listbox" Or arrObj(0) = "listview"  Then
				object.ExtendSelect CInt(arrKeyValue(1)) 
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			 Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If			 	  	
		Case "maximize"		
			If arrObj(0) = "window" Or arrObj(0) = "dialog"  Or arrObj(0) = "popupwindow" Then
				If curParent.GetROProperty ("enabled") Then 'Check to see if a window is able to be maximized.
					If curParent.GetROProperty ("maximizable") Then
						curParent.Maximize
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					End If
				End If
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If	
		Case "minimize"
			If arrObj(0) = "window" Or arrObj(0) = "dialog"  Or arrObj(0) = "popupwindow" Then
				If curParent.GetROProperty ("enabled") Then 'Check to see if a window is able to be minimized.
					If curParent.GetROProperty ("minimizable") Then
						curParent.Minimize
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					End If
				End If
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If					
		Case "next"
			If arrObj(0) = "spinner"  Then
				object.Next
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If
		Case "nextline"
			If arrObj(0) = "scrollbar"  Then
				If UBound(arrKeyValue) > 0 Then
					object.NextLine CInt(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Else
					object.NextLine
				End If
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "nextpage"
			If arrObj(0) = "scrollbar"  Then
				If UBound(arrKeyValue) > 0 Then
					object.NextPage CInt(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Else
					object.NextPage
				End If
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "press"
			If arrObj(0) = "toolbar"  Then
				wait 2
				object.Press arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "previous"
			If arrObj(0) = "spinner"  Then
				object.Prev
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If
      	Case "prevpage"
			If arrObj(0) = "scrollbar"  Then
				If UBound(arrKeyValue) > 0 Then
					object.PrevPage CInt(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Else
					object.Prevpage
				End If
			 Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If		
		Case "prevline"
			If arrObj(0) = "scrollbar"  Then
				If UBound(arrKeyValue) > 0 Then
					object.PrevLine CInt(arrKeyValue(1))
					Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
				Else
					object.PrevLine
				End If
			 Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If					
		Case "restore"
			If arrObj(0) = "window" Or arrObj(0) = "dialog"  Or arrObj(0) = "popupwindow" Then
				curParent.Restore
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			 Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If					
		Case "showdropdown"
			If arrObj(0) = "toolbar"  Then
				object.ShowDropdown arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If		
		Case "set" 'If condition included to use 'set' also for a radiobutton
			If arrObj(0) = "scrollbar"  Or arrObj(0) = "textbox" Or arrObj(0) = "combobox" or arrObj(0) ="wintextbox" or arrObj(0) ="spinner"  Then
				object.Set arrKeyValue(1)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			ElseIf  arrObj(0) = "radiobutton" Then
				object.Set
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "select"
			If arrObj(0) = "toolbar"  Or arrObj(0) = "treeview" Or arrObj(0) = "listview" Or arrObj(0) = "menu" Or arrObj(0) = "combobox"  Or arrObj(0) = "listbox" Or arrObj(0) = "tab" Then
				object.Select arrKeyValue(1)  
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				  		
		Case "selectindex"
			If  arrObj(0) = "listview"  Or arrObj(0) = "combobox"  Or arrObj(0) = "listbox" Or arrObj(0) = "tab"Then
				object.Select CInt(arrKeyValue(1))
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "selectrange"
			If  arrObj(0) = "listview"  Or arrObj(0) = "listbox" Then
				object.SelectRange arrKeyIndex(1),arrKeyIndex(2)
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "selectrangeindex"
			If  arrObj(0) = "listview"  Or arrObj(0) = "listbox" Then
				object.SelectRange CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2))
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If				
		Case "setselection"
			If arrObj(0) = "editor" Then
				object.SetSelection CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2)),CInt(arrKeyIndex(3)),CInt(arrKeyIndex(4))' for editor
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			ElseIf arrObj(0) = "textbox" Then
				object.SetSelection CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2))'for textbox
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		   	Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If	  	
		Case "setcaretpos"
			If arrObj(0) = "editor" Then
				object.SetCaretPos CInt(arrKeyIndex(1)),CInt(arrKeyIndex(2))'for editor
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			ElseIf arrObj(0) = "textbox" Then
				object.SetCaretPos CInt(arrKeyIndex(1))'for textbox
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		   	Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If			
		Case "setdate" 
			If arrObj(0) = "calendar" Then
				Select Case LCase(arrKeyValue(1)) 'setting operation for calendar object
					Case "now"
						object.SetDate Now
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(1)&" - " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					Case "date"
						object.SetDate Date
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(1)&" - " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					Case Else
						If IsDate(arrKeyValue(1)) Then
							object.SetDate CDate(arrKeyValue(1))
						Else
							Reporter.ReportEvent micFail,"Select Date","Invalid date provided."
						End If
				End Select
		   	Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If					
		Case "settime"
			If arrObj(0) = "calendar" Then
				Select Case LCase(arrKeyValue(1)) 'setting current time of the system
					Case "now"
						Dim curTime
						curTime = Time()
						object.SetTime curTime
						Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(1)&" - " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
					Case Else
						If IsDate(arrKeyValue(1)) Then
							object.SetTime arrKeyValue(1)
						Else
							Reporter.ReportEvent micFail,"Select Date","Invalid Time provided."
						End If
				End Select 	
		   	Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If					
		Case "type"
			object.Type arrKeyValue(1)
			Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		Case "textclick"
              Call Func_SelectText(arrKeyIndex(1))			
			  Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
		Case "uncheck"
			If arrObj(0) = "treeview" Or arrObj(0) = "listview" Then
				object.SetItemState arrKeyValue(1),micUnChecked
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			ElseIf  arrObj(0) = "checkbox" Then
				object.Set "OFF"
				object.WaitProperty "checked",False,3000
				Reporter.ReportEvent micDone, arrObj(0) &"  - "&arrKeyValue(0), "Action " &arrKeyValue(0) &" performed successfully on "&arrObj(0)
			 Else
				Reporter.reportevent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
			End If			  	
       Case Else
			If (LCase(Trim(arrObj(0)))= "sqlexecute") Or (LCase(Trim(arrObj(0)))= "sqlvaluecapture") Or (LCase(Trim(arrObj(0)))= "sqlcheckpoint") Or (LCase(Trim(arrObj(0)))= "sqlmultiplecapture") Then
'				NOTE: The following variables should be present in the Environment_values.xml File attached to the script
				Environment("connectionString") = "DRIVER={" & Environment("dbDRIVER") & "};UID=" & Environment("dbUID") & ";PWD=" & Environment("dbPWD") & ";SERVER=" & Environment("dbServer") & "_" & Environment("dbHost") & ".world"
				Select Case LCase(Trim(arrObj(0))) 'to perform sql operation
					Case "sqlexecute"
						strSQL = arrObj(1)
						For inti = 0 to (Func_RegExpMatch ("##\w*##",arrObj(1),aPosition,aMatch) - 1)
							strReplace = "'" & Environment.Value(Replace(aMatch(inti),"##","",1,-1,1)) & "'"
							strSQL = Func_gfRegExpReplace(aMatch(inti), strSQL, strReplace)
						Next								
						Set dbConn = CreateObject("ADODB.Connection")	'Create the DB Object
						dbConn.Open Environment.Value("connectionString")
						Set dbRs = dbConn.Execute(strSQL)	'Execute the query
						If LCase(Trim(arrKeyIndex(0)))="commit" Then
							If LCase(Trim(arrKeyIndex(1)))="yes" Then
								Set dbRs = dbConn.Execute("Commit")	'Commit the database
							Else
								Reporter.ReportEvent micDone,"Keyword Check at Line no - " & intRowCount & vbcrlf & "'" & arrObj(0) & "'","The changes made are not commited to the database"
							End If
						Else
							Reporter.ReportEvent micFail, "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrKeyIndex(0) & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."  
						End If
						dbConn.Close	'Close the database connection
						Set dbConn = Nothing
						Reporter.ReportEvent micDone, " Sql Operation", "Query executed successfully" 
					Case "sqlvaluecapture"
						strSQL = arrObj(1)
						For inti = 0 to (Func_RegExpMatch ("##\w*##",arrObj(1),aPosition,aMatch) - 1)
							strReplace = "'" & Environment.Value(Replace(aMatch(inti),"##","",1,-1,1)) & "'"
							strSQL = Func_gfRegExpReplace(aMatch(inti), strSQL, strReplace)
						Next
						Environment.Value(arrKeyIndex(0)) = Func_gfQuery(strSQL)
						Reporter.ReportEvent micDone, " Sql value capture", "Sql value captured successfully" 
					Case "sqlcheckpoint"
                        DataTable.GetSheet("Action1").SetCurrentRow(1)
                        DbTable(arrKeyIndex(0)).SetTOProperty "connectionstring", Environment.Value("connectionString")	'Set the TO property of connection string.
						strSQL = arrAction(1)
						For inti = 0 to (Func_RegExpMatch ("##\w*##", arrAction(1),aPosition,aMatch) - 1)
							strReplace = "'" & Environment.Value(Replace(aMatch(inti),"##","",1,-1,1)) & "'"
							strSQL = Func_gfRegExpReplace(aMatch(inti), strSQL, strReplace)
						Next
						DbTable(arrKeyIndex(0)).SetTOProperty "source", strSQL	'Change the DB OBjects source(SQL) statement.
						DbTable(arrKeyIndex(0)).Check CheckPoint(arrKeyIndex(0)) 'Execute the DB checkpoint.
						Reporter.ReportEvent micDone, " sql check point", "sql check point successfully" 
					Case "sqlmultiplecapture"
                         DataTable.GetSheet("Action1").SetCurrentRow(1)
                        DbTable(arrKeyIndex(0)).SetTOProperty "connectionstring", Environment.Value("connectionString")	'Set the TO property of connection string.
						strSQL = arrAction(1)
						For inti = 0 to (Func_RegExpMatch ("##\w*##", arrAction(1),aPosition,aMatch) - 1)
							strReplace = "'" & Environment.Value(Replace(aMatch(inti),"##","",1,-1,1)) & "'"
							strSQL = Func_gfRegExpReplace(aMatch(inti), strSQL, strReplace)
						Next
						DbTable(arrKeyIndex(0)).SetTOProperty "source", strSQL	'Change the DB OBjects source(SQL) statement.
						DbTable(arrKeyIndex(0)).Output CheckPoint(arrKeyIndex(0)) 'Execute the DB output checkpoint.
						Reporter.ReportEvent micDone, " sql multiple capture", "sql multiple captured successfully" 
                    End Select
			Else
				If ((arrObj(0) <> "split") And (arrObj(0) <> "random")) Then
					Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & arrObj(0) & "'  not supported." , "Please verify Keyword entered."
				If htmlreport = "1"  Then
					Call Update_log(MAIN_FOLDER, g_sFileName, "fail")' calling function update log to create an execution log in HTML file-
				End If
				End If	
			End If 
	End Select 
	Dim intNum      'This variable stores the value in arrObj(1)
	Dim strvar      'This variable is used to store the string value in 4th Column of Datasheet
	Dim strsplit    'This variable is used to store the array after Split operation
	Dim strlen      'This variable is used to store the string value in 5th Column of Datasheet
	Dim strstore1    'This variable is used to store the elements present in the fourth Column 
	Dim arrVals		'This variabler is used store the split array elements
	Dim strvarstore'This variable is used to store the split element 	
	Dim intval
	Select Case LCase(arrObj(0)) 'to perform special/common operation
		Case "random"
            Randomize   
			intNum = arrObj(1)
			strvar = Cstr(Trim(DataTable.Value(4,dtLocalSheet)))
			Environment.Value(strvar) = intNum * Rnd
		Case "split"
			strvar = Split(arrObj(1),"^",-1,1)
			For inti= 0 to Ubound(strvar)
				If(Instr(1,strvar(inti),"#") = 1) Then
					strvar(inti) = Environment.Value(Right(strvar(inti),Len(strvar(inti))-1))
				End If
			Next
			strsplit = Split(strvar(0),strvar(1),-1,1)
			If  DataTable.Value(5,dtLocalSheet) <> "" Then
				strlen=Cstr(Trim(DataTable.Value(5,dtLocalSheet)))
				 Environment.Value(strlen)=Ubound(strsplit)
			End If
			strstore1 = Cstr(Trim(DataTable.Value(4,dtLocalSheet)))
			arrVals = Split(strstore1,";",-1,1)
			For inti = 0 to Ubound(arrVals)
				strvarstore  = Split(arrVals(inti),":",2,1)
				intval = Cint(strvarstore(1))
				Environment.Value(strvarstore(0)) = strsplit(intval)
			Next			   
	End Select 'End of special/common operations  

				If   Environment.Value ("iperform") =1 Then
					Call Func_CaptureScreenshot("test",intRowCount) 'call to screencapture function to take screenshot
					wait 1
				End If
End Function'End of perform
''#################################################################################################

'#################################################################################################
'Function name 		: Func_Store_Dotnet
'Description    	: If User requires to store any property of a particular object into a variable 
'					  then this fuction can be used.
'Parameters     	: The Object details as the full hierarchical Path of the Object goes as parameter 
'			          to the function.            
'Assumptions    	: None 
'#################################################################################################
'The following function is for StoreValue keyword.
'#################################################################################################
Function Func_Store_DotNet(object,arrObj)
	Dim strPropName  'Stores the name of the Property to be stored
	Dim arrPropSplit 'Stores the Property name and Variable Name.
	Dim intGRowNum	'Stores the row number
	Dim intGColNum	'Stores the Column number

	If  CInt(InStr(1, arrObj(0),"d_")) > 0 Then
		VarName = DataTable.Value(4,dtLocalsheet)
		 Select Case LCase(arrObj(0))  'for performing time and date operation
			Case "d_currenttime"
				Environment.Value(VarName) = FormatDateTime(Now(),4)
			Case "d_currentdate"
				Environment.Value(VarName) = FormatDateTime(Now,2)
				propSplit = Split(Environment.Value(VarName),"/",-1,1)
				If propSplit(0) < 10 Then
					propSplit(0) = 0 & propSplit(0) 
					Environment.Value(VarName) = propSplit(0) & "/" & propSplit(1) & "/" & propSplit(2)
				End If
		End Select
	Else
		arrPropSplit = Split(DataTable.Value(4,dtLocalsheet),":",-1,1) 'splitting the value in the 4th Column into Property and Variable Name
		strPropName = arrPropSplit(0)		  'Storing the Property into strPropName 
		VarName = arrPropSplit(1)			  'Storing the Variable Name into VarName 
		Select Case LCase(strPropName)		  'Case to store the required property of the variable into the VarName variable 
			Case "itemscount"
				Select Case LCase(arrObj(0)) 'getting the itemscount of the objects
					Case "toolbar"
						Environment.Value(VarName) = object.GetItemsCount
					Case "treeview"
						Environment.Value(VarName) = object.GetItemsCount
					Case Else
						Environment.Value(VarName) = object.GetROProperty("items count")
				End Select
			Case "enabled"
				Environment.Value(VarName) = Not(CBOOL(object.GetROProperty("disabled")))
			Case "columncount"
				Environment.Value(VarName) = object.ColumnCount(1)
			Case "rowcount" 
				Environment.Value(VarName) = object.RowCount
			Case "filename"
				Environment.Value(VarName) = object.GetROProperty("file name")
			Case "imagetype"
				Environment.Value(VarName) = object.GetROProperty("image type")
			Case "defaultvalue"
				Environment.Value(VarName) = object.GetROProperty("default value")
			Case "maxlength"
				Environment.Value(VarName) = object.GetROProperty("max length")
			Case "allitems"
				Select Case LCase(arrObj(0))
					Case "toolbar"
						Environment.Value(VarName) = object.GetContent
					Case "treeview"
						Environment.Value(VarName) = object.GetContent
					Case Else
						Environment.Value(VarName) = object.GetROProperty("all items")
				End Select
			Case "selectiontype"
				Environment.Value(VarName) = object.GetROProperty("select type")
			Case "exist"
				If arrObj(0) = "window" Or arrObj(0) = "dialog" Then
					Environment.Value(VarName) = Cstr((curParent.Exist(5)))
				Else
					Environment.Value(VarName) = Cstr((object.Exist(5)))
				End If
			Case "itemexist"
				 Select Case LCase(arrObj(0))
					Case "toolbar"
						Environment.Value(VarName) = object.ItemExists
				End Select
			Case "selection"
				Select Case LCase(arrObj(0))
					Case "toolbar"
						Environment.Value(VarName) = object.GetSelection
					Case "treeview"
						Environment.Value(VarName) = object.GetSelection
					Case Else
						Environment.Value(VarName) = object.GetROProperty("selection")
				End Select
			Case "selectioncount"
				Environment.Value(VarName) = object.GetROProperty("selected items count")
			Case "getcelldata"
				If arrObj(0)="table" Then
					intGRowNum = Cint(arrPropSplit(2))
					intGColNum = Cint(arrPropSplit(3))
					Environment.Value(VarName) = object.GetCellData(intGRowNum,intGColNum)
				Else
					Reporter.reportevent micFail,  "Keyword Check at Line no - " & intRowCount, "Keyword - '" & strPropName & "'  not supported for -" &arrObj(0) , "Please verify Keyword entered."
				End If
			Case "tablesearch"
				Call Func_tablesearch(object,arrTableindex(1))
			Case Else
				If arrObj(0) = "window" Or arrObj(0) = "dialog" Then
					Environment.Value(VarName) = curParent.GetROProperty(strPropName)
				Else
					Environment.Value(VarName) = object.GetROProperty(strPropName)
				End If
		End Select
	End If

    If cint(Introwcount)<=cint(Environment("intEndRow")) Then
		If Cint(Environment("intStartRow"))<=Cint(Introwcount) Then
			Call DebugGetEnv()      'to call debug function for execution log status in HTML file
        End If 
    End If

End Function
'#################################################################################################

'#################################################################################################
'Function name 	: Func_Check_Dotnet
'Description   	: This function is used for all the checking operations to be performed on the AUT.
'Parameters    	: The Object details on which check has to be performed along with details of the 
'		  	fourth Column of the current row in local sheet and the current row number in the 
'	          	local Sheet.           
'Assumptions	: NA
'#################################################################################################
'The following function is for 'Check' keyword.
'#################################################################################################
Function Func_Check_DotNet(object,arrAction,arrKeyValue,arrKeyIndex,intRowCount)
   Dim iStatus  'Stores the final status of the check (pass, fail, done,etc.)
   Dim actualValue  'stores the actual property value of the object at run time
   Dim expectedValue 'stores the value of the property of the object defined in datatable
   Dim strStatus ' 'stores the status of the execution
   Dim strStatus1 'Stores a particular string based on the actual/expected property value during the check operation. This string is used in the report statement
   Dim reportStep ''stores the step to be performed on the object
   Dim reportStepPass   'stores the pass result of the object 
   Dim reportStepFail   'stores the fail result of the object
   Dim itemFound   ''initialization variable
   Dim actualCount  'stores the run time property of an object for itemcount
   Dim checkval  ' initialization variable used for func_checkvaluereturn
   checkval = 2
	iStatus = "iDone"

  Select Case LCase (arrKeyIndex(0))  'checking operation action for the object
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
				If LCase(arrKeyIndex(1)) = "true" Or LCase(arrKeyIndex(1)) = "false" Then
				ActualValue = CBool(object.GetROProperty(LCase(arrKeyIndex(0))))
				ExpectedValue = CBool(arrKeyIndex(1))
					If LCase(ExpectedValue) = "false" Then
					strStatus = " not "
					Else
					strStatus = ""
					End If
			reportStep = "Verify that '" &  arrAction(1)  & "'  "&  arrAction(0) &" is "&  strStatus &" " &arrKeyIndex(0) & "."
               If ExpectedValue = ActualValue Then
					reportStepPass = "The '"&  arrAction(1)  &"'  "& arrAction(0) &" is " & strStatus & " " & arrKeyIndex(0) & ", which is as expected."
					iStatus = "iPass"
					keyword = 0
					checkval=1
				Else
					If LCase(ExpectedValue) = "false" Then
						reportStepFail = "The '"&  arrAction(1)  &"'  "&  arrAction(0) &" is " & arrKeyIndex(0) & ", which is not as expected."
					Else
						reportStepFail = "The '"&  arrAction(1)  &"'  "&  arrAction(0) &" is  not " & arrKeyIndex(0) & ", which is not as expected."	
					End If
					iStatus = "iFail"
					keyword = 1
					checkval=0
				End If	   		 
		End If  
		
			Case "focused"
				If LCase(arrKeyValue(1)) = "true" Or LCase(arrKeyIndex(1)) = "false" Then
				ActualValue = CBool(object.GetROProperty(LCase(arrKeyIndex(0))))
				ExpectedValue = CBool(arrKeyIndex(1))
					If LCase(ExpectedValue) = "false" Then
					strStatus = " not "
					Else
					strStatus = ""
					End If
			reportStep = "Verify that '" & arrAction(1)  & "'  "& arrAction(0) &" is "&  strStatus &" " &arrKeyIndex(0) & "."			
               If ExpectedValue = ActualValue Then
					reportStepPass = "The '"& arrAction(1)  &"'  "& arrAction(0) &" is " & strStatus & " " & arrKeyIndex(0) & ", which is as expected."
					iStatus = "iPass"
					keyword = 0
					checkval=1
				Else
					If LCase(ExpectedValue) = "false" Then
						reportStepFail = "The '"& arrAction(1)  &"'  "& arrAction(0) &" is " & arrKeyIndex(0) & ", which is not as expected."
					Else
						reportStepFail = "The '"& arrAction(1)  &"'  "&arrAction(0) &" is  not " & arrKeyIndex(0) & ", which is not as expected."	
					End If
					iStatus = "iFail"
					keyword = 1
					checkval=0
				End If
		End If	

		
			Case "visible"
				If LCase(arrKeyValue(1)) = "true" Or LCase(arrKeyIndex(1)) = "false" Then
					ActualValue = CBool(object.GetROProperty(LCase(arrKeyIndex(0))))
					ExpectedValue = CBool(arrKeyIndex(1))
					If LCase(ExpectedValue) = "false" Then
						strStatus = " not "
					Else
						strStatus = ""
					End If
						reportStep = "Verify that '" & arrAction(1) & "'  "& arrAction(0) &" is "&  strStatus &" " &arrKeyIndex(0) & "."			
               If ExpectedValue = ActualValue Then
					reportStepPass = "The '"& arrAction(1) &"'  "& arrAction(0) &" is " & strStatus & " " & arrKeyIndex(0) & ", which is as expected."
					iStatus = "iPass"
					keyword = 0
					checkval=1
				Else
					If LCase(ExpectedValue) = "false" Then
						reportStepFail = "The '"& arrAction(1) &"'  "& arrAction(0) &" is " & arrKeyIndex(0) & ", which is not as expected."
					Else
						reportStepFail = "The '"& arrAction(1)  &"'  "& arrAction(0) &" is  not " & arrKeyIndex(0) & ", which is not as expected."	
					End If
					iStatus = "iFail"
					keyword = 1
					checkval=0
				End If
		End If	

		Case "itemcount"
				expectedValue = Cdbl(arrKeyIndex(1))
				actualValue  = CInt(object.GetROProperty("items count"))
				reportStep = "Verify the number of items in  '"& arrAction(1) &"'  "& arrAction(0) &" is "& arrKeyIndex(1) & "."
			If Lcase(expectedValue)=Lcase(actualValue) Then
				reportStepPass = "The number of items in '"& arrAction(1) &"'  "& arrAction(0) &" is "& actualValue & ", which is as expected."
				iStatus="iPass"
				keyword = 0
				checkval = 1
				Else
				reportStepFail =  "The number of items in '"& arrAction(1) &"'  "& arrAction(0) &" is "& actualValue & ", which is not as expected of '"&arrKeyIndex(1)&"' ."
				iStatus="iFail"
				keyword = 1
				checkval = 0
			End If
            
		Case "columncount"
				expectedValue = CInt(arrKeyIndex(1))
				actualValue  = CInt(object.ColumnCount)
				reportStep = "Verify the number of columns in  '"& arrAction(1) &"'  "& arrAction(0) &" is "& arrKeyIndex(1) & "."
			If Lcase(expectedValue)=Lcase(actualValue) Then
				reportStepPass = "The number of columns in '"& arrAction(1) &"'  "& arrAction(0) &" is "& actualValue & ", which is as expected."
				iStatus="iPass"
				keyword = 0
				checkval = 1
			Else
				reportStepFail =  "The number of columns in '"& arrAction(1) &"'  "& arrAction(0) &" is "& actualValue & ", which is not as expected of '"&arrKeyIndex(1)&"' ."
				iStatus="iFail"
				keyword = 1
				checkval = 0
			End If
		
		Case "rowcount"
				expectedValue = CInt(arrKeyIndex(1))
				actualValue  = CInt(object.RowCount)
				reportStep = "Verify the number of rows in  '"& arrAction(1) &"'  "& arrAction(0) &" is "& arrKeyIndex(1) & "."
			If Lcase(expectedValue)=Lcase(actualValue) Then
				reportStepPass = "The number of rows in '"& arrAction(1) &"'  "& arrAction(0) &" is "& actualValue & ", which is as expected."
				iStatus="iPass"
				keyword=0
				checkval=1
			Else
				reportStepFail =  "The number of rows in '"& arrAction(1) &"'  "& arrAction(0) &" is "& actualValue & ", which is not as expected of '"&arrKeyIndex(1)&"'."
				iStatus="iFail"
				keyword=1
				checkval=0
			End If			
			
		Case "text"
             expectedValue = Trim(arrKeyIndex(1))
			actualValue  = Trim(object.GetROProperty("text"))
			reportStep = "Verify that  "& arrKeyIndex(0) & "-'" & arrKeyIndex(1) &"' is displayed in '"& arrAction(1) &"'  "& arrAction(0) & "."
			If Lcase(expectedValue)=Lcase(actualValue) Then
				reportStepPass = "The text displayed in '"& arrAction(1) &"'  "& arrAction(0) &" is '"& actualValue & "', which is as expected."
				keyword=0
				checkval=1
			Else
				reportStepFail =  "The text displayed in '"& arrAction(1) &"'  "& arrAction(0) &" is '"& actualValue & "', which is not as expected."
				keyword=1
				checkval=0
			End If
			
		Case "selection"
			expectedValue = Trim(arrKeyIndex(1))
			actualValue  = Trim(object.GetROProperty("selection"))
			reportStep = "Verify that  "& arrKeyIndex(1) & " is selected in '"& arrAction(1) &"'  "& arrAction(0) & ""
			If Lcase(expectedValue)=Lcase(actualValue) Then
				reportStepPass = "The selected item in '"& arrAction(1) &"'  "& arrAction(0) &" is "& actualValue & ", which is as expected."
				keyword=0
				checkval=1
			Else
				reportStepFail =  "The selected item in '"& arrAction(1) &"'  "& arrAction(0) &" is "& actualValue & ", which is not as expected of '"&arrKeyIndex(1)&"' ."
				keyword=1
				checkval=0
			End If
			
		Case "exist"
			If LCase(arrKeyIndex(1)) = "true" Or LCase(arrKeyIndex(1)) = "false" Then
				actualValue = CStr(object.Exist(10))
				expectedValue = CStr(arrKeyIndex(1))
				If LCase(expectedValue) = "false" Then
					strStatus = " does not exist "
				Else
					strStatus = " exists"
				End If
				If LCase(actualValue)="false" Then
					strStatus1="does not exist"
				Else
					strStatus1="exists"
				End If
				reportStep =  "Verify that '" & arrAction(1) & "'  "& arrAction(0) &  strStatus & "."
				If Lcase(actualValue)=Lcase(expectedValue) Then
				reportStepPass = "The "& arrAction(0) &" '"& arrAction(1) &"' " &  strStatus1 & " , which is as expected."
				keyword=0
				checkval=1
				else
				reportStepFail = "The "& arrAction(0) &" '"& arrAction(1) &"' " &  strStatus1 & " , which is not as expected."
				keyword=1
				checkval=0
				End If                			
			End If
			
		Case "checked"
			If LCase(arrKeyIndex(1)) = "true" Or LCase(arrKeyIndex(1)) = "false" Then
				actualValue = object.GetROProperty("checked")
				expectedValue = UCase(arrKeyIndex(1))
		 		If LCase(actualValue) = "false"  Then
					strStatus = " unchecked "
				Else
					strStatus = "checked"
				End If
				If LCase(expectedValue) = "false"  Then
					strStatus1 = " unchecked "
				Else
					strStatus1 = "checked"
				End If
				reportStep =  "Verify that '" & arrAction(1) & "'  "& arrAction(0) &" is "&  strStatus1 & "."
				If LCase(actualValue)=LCase(expectedValue) Then
					reportStepPass = "The  "&arrAction(0) & " " &arrAction(1) &" is "&  strStatus & ", which is as expected."
					iStatus="iPass"
					keyword=0
					checkval=1
				Else
					reportStepFail =  "The "& arrAction(0) & " " &arrAction(1) &" is "&  strStatus & ", which is not as expected."
					iStatus="iFail"
					keyword=1
					checkval=0
				End If
			End If
			
		Case "tabexist"
           	itemFound = 2
			actualCount = object.GetROProperty("items count")
			For inti = 0 to actualCount-1
				actualValue = object.GetItem(inti)
				If Trim(actualValue) = Trim(arrKeyIndex(1)) Then
					itemFound =0
					Exit For
				End If
			Next
			reportStep =  "Verify that Tab item '" & arrKeyIndex(1) & "' is present in the Tab '" & arrAction(0) &"."
			If itemFound <> 0 Then
				reportStepFail = "The Tab item  '"&arrKeyIndex(1)&"' does not exist in the Tab '" & arrAction(0) & "."
				iStatus = "iFail"
				keyword = 1
				checkval=0
			Else
				reportStepPass = "The Tab item  '"&arrKeyIndex(1)&"' exists in the Tab '" & arrAction(0) & "."
				iStatus = "iPass"
				keyword = 0 
				checkval=1
			End If
			
		Case "tabnotexist"
			itemFound = 2
			actualCount = object.GetROProperty("items count")
				For inti = 0 to actualCount-1
					actualValue = object.GetItem(inti)
				    If Trim(actualValue) = Trim(arrKeyIndex(1)) Then
						 itemFound =0
						 Exit For
					End If
				Next
				reportStep =  "Verify that Tab item '" & arrKeyIndex(1) & "' is not present in the Tab '" & arrAction(0) &"."
				If itemFound <> 0 Then
					reportStepPass = "The Tab item  '"&arrKeyIndex(1)&"' does not exist in the Tab '" & arrAction(0) & "."
				 	iStatus = "iPass"
					keyword = 0
					checkval=1
				Else
					reportStepFail = "The Tab item  '"&arrKeyIndex(1)&"' exists in the Tab '" & arrAction(0) & "."
					iStatus = "iFail"
					keyword = 1
					checkval=0					
				End If
				
		Case "itemexist"
			itemFound = 2
			reportStep = "Verify that  in " & arrAction(0) &" - '" & arrAction(1)& "', '" & arrKeyIndex(1) & "' is present."
			actualCount = CInt(object.GetItemsCount)
			For inti = 0 to actualCount -1
				actualValue = object.GetItem(inti)
				If actualValue = arrKeyIndex(1)Then
					itemFound = 0
					Exit For
				End If
			Next
			If itemFound <> 0 Then
				reportStepFail = "The item  '"&arrKeyIndex(1)&"' does not exist in " & arrAction(0) & "."
				iStatus = "iFail"
				keyword = 1
				checkval=0
			Else
				reportStepPass = "The item  '"&arrKeyIndex(1)&"' exists in " & arrAction(0) & "-'" & arrAction(1) & "'."
				iStatus = "iPass"
				keyword = 0
				checkval=1
			End If
			
    		Case "windowtext"
				Dim l, t, r, b
				Dim strText 
				l = 2
				t = 2
				r = CInt(object.GetROProperty("width")) - 10
				b = CInt(object.GetROProperty("height")) - 10
				 strText = object.GetVisibleText(l,t,r,b)    	
				strText = Replace(strText,Chr(13),"",1,-1,1)	 'Replace the Carriage Return Character   			
				strText = Replace(strText,Chr(10),"",1,-1,1) 'Replace the New Line Character   				
				strText = Replace(strText," ","",1,-1,1)'Remove any spaces.
				'Need to handle an optional True/False parameter.	
				expectedValue = True'Assume True if the optional parameter is not supplied.
					If UBound(arrKeyIndex) >1 Then
					ExpectedValue = CBool(arrKeyIndex(2))
				End If
				If ExpectedValue = false Then
					strStatus = " is not "
				else
					strStatus = " is "	
				End If
				ActualValue =  Func_gfRegExpTest(arrKeyIndex(1), strText)
				reportStep= "Verify that Text:'" & arrKeyIndex(1) & "' ,"& strStatus &" displayed in the Window:'" & arrAction(1) &"'."
				If LCase(ActualValue)=LCase(ExpectedValue) Then
					reportStepPass= "Text:'" & arrKeyIndex(1) & vbCrLf & "Exist:" & ActualValue & vbCrLf & "Actual:" & strText  &", which is as expected"
					iStatus="iPass"
					keyword=0
					checkval=1
					Else if LCase(ExpectedValue)="false" Then
					reportStepFail= "Text:'" & arrKeyIndex(1) & vbCrLf & "Exist:" & ActualValue & vbCrLf & "Actual:" & strText &", which is not as expected"
					Else
					reportStepFail= "Text:'" & arrKeyIndex(1) & vbCrLf & "Not Exist:" & ExpectedValue & vbCrLf & "Actual:" & strText &", which is not as expected"
                	End If
					iStatus="iFail"
					keyword=1
					checkval=0
			End If

			Case "itemspresent"
         		Dim iCount1
         		Dim iCount2  'stores the value of itemcount property
                Dim sItem
         		Dim ArrItems
				reportStep="Check the Items in the '"&arrAction(1)&"' " & arrAction(0)
				iCount1=Ubound(arrKeyIndex)
				iCount2=object.GetItemsCount
                For intj = 1 to iCount1
					iStatus = "inew"
                    For inti = 0 to iCount2-1
						sItem = object.GetItem(inti)
						If Trim(sItem) = Trim(arrKeyIndex(intj)) Then
							Reporter.ReportEvent micPass, reportStep, "The item: "& arrKeyIndex(intj) &" exists  in the '"&arrAction(1)&"'  " &arrAction(0)
							iStatus = "iPass"
							keyword=0
							checkval=1
                            Exit For
						Else
						Reporter.ReportEvent micFail, reportStep, "The item: "& arrKeyIndex(intj) &" does not exist in the '"&arrAction(1)&"' " & arrAction(0)
						iStatus="iFail"
						keyword=1
						checkval=0
                        End If
					Next
				Next

		Case "table"
			actualValue = object.GetROProperty("currentrowindex")
			If actualValue = -1 Then
				strStatus = "empty"
				actualValue = strStatus
			ElseIf actualValue = 0 Then
				strStatus = "not empty"
				actualValue = strStatus
			End If
				expectedValue = LCase(arrKeyIndex(1))
                 reportStep =  "Verify that the "& arrAction(0) &"  '" &arrAction(1)& "' is "&expectedValue& "." 
				 If LCase(actualValue)=LCase(expectedValue) Then
					reportStepPass =  "The "& arrAction(0) &"  '"&arrAction(1)&"' is "&strStatus& ", which is as expected."
					iStatus="iPass"
					keyword=0
					checkval=1
				Else
					reportStepFail =  "The "& arrAction(0) &"  '"&arrAction(1)&"' is "&strStatus& ", which is not as expected."
					iStatus="iFail"
					keyword=1
					checkval=0
				 End If               

		Case Else
			actualValue = CStr(object.GetROProperty(arrKeyIndex(0))) 
			expectedValue = (arrKeyIndex(1)) 
			reportStep =  "Verify the "&arrKeyIndex(0)&" property of '" & arrAction(1) & "'  "& arrAction(0) &"." 
			If Lcase( actualValue)=Lcase(expectedValue)Then
					reportStepPass = "The property " &arrKeyIndex(0)&" of '" & arrAction(1) & "'  "& arrAction(0) &" is "&actualValue&", which is as expected."
					istatus="iPass"
					keyword=0
					checkval=1
			Elseif Lcase(actualValue)="" Then 'checking for unavailable properties of an object
					reportStep =  "Verify the "&arrKeyIndex(0)&" property of '" & arrAction(1) & "'  "& arrAction(0) &"." 
					reportStepFail = "Property does not exists. Please verify property entered."
					Call Func_CaptureScreenshot("checkpoint",intRowCount)	    'call to screencapture function to take screenshot
			Exit Function
			Else
					reportStepFail = "The property " &arrKeyIndex(0)&" of '" & arrAction(1) & "'  "& arrAction(0) &" is "&actualValue&", which is not as expected value of "&expectedValue&"."
					istatus="iFail"
					keyword=1
					checkval=0
   			End if			
        End Select      				

		'Based on the status of the Check operations performed above, Status is set and Results reporting is done
		Select Case iStatus
			Case "iDone"
				If  arrKeyValue(0)<>"text" Then
					actualValue=LCase(actualValue)
					expectedValue=LCase(expectedValue)
				End If
				If actualValue = expectedValue Then
					Reporter.ReportEvent micPass,reportStep,reportStepPass
					iStatus = "iPass"                 
				Else
					Reporter.ReportEvent micFail,reportStep,reportStepFail
					Call Func_CaptureScreenshot("checkpoint",intRowCount)	   'call to screencapture function to take screenshot
					iStatus = "iFail"                        
				End If
			Case "iPass"
				Reporter.ReportEvent micPass,reportStep,reportStepPass
				If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkpass")	  ' calling function update log to create an execution log in HTML file
				End If	
			Case "iFail"
				Reporter.ReportEvent micFail,reportStep,reportStepFail
				If htmlreport = "1" Then
					 Call Update_log(MAIN_FOLDER, g_sFileName, "checkfail")	   ' calling function update log to create an execution log in HTML file
				End If
				Call Func_CaptureScreenshot("checkpoint",intRowCount)	    'call to screencapture function to take screenshot
		End Select
		
		 If Err.Number <> 0 Then
			 Reporter.ReportEvent micFail,"ERROR -Occurred at Line :  "&DataTable.LocalSheet.GetCurrentRow,  "Error Description : "& Err.Description
			 Call Func_CaptureScreenshot("checkpoint",intRowCount)	   'call to screencapture function to take screenshot
			 Err.Clear
		End If

			 If Datatable.Value(5,dtLocalSheet) <> empty Then 
				Call Func_CheckValueReturn(intRowCount,checkval)
						 If cint(Introwcount)<=cint(Environment("intEndRow")) Then
							If Cint(Environment("intStartRow"))<=Cint(Introwcount) Then
							  Call DebugGetEnv()       'to call debug function for execution log status in HTML file    
							End If	
						End If
			End If

					If   Environment.Value("icheck")=1  Then
						Call Func_CaptureScreenshot("test",intRowCount)      'call to screencapture function to take screenshot
						wait 1
					End If
		
End Function
'#################################################################################################
'
''####################################################################################################
'Function name 	: Func_getRowNum
'Description        : If the user wants to retrieve the row number in which specified text is present in table, this function can be used
'Parameters       	: The object(i.e, Table)  in which the search operation needs to be performed
'					  The text to be searched in the table and the number of columns in the table.
'Return Value		: This function returns row number of the given celltext
'Assumptions     	: NA
'#####################################################################################################
'The following function is used for rownum Keyword
'#####################################################################################################
Function Func_getRowNum_Dotnet(object,strSearch,strReturnVal)
	Dim arrCol				
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
	If (intCheck = -1) Then
        If (initial ="check") Then
	        Reporter.ReportEvent micPass,"Cell text "&strRowVal&"  should not be present in the table","Cell text is not present in the table"
        End If
	 Else
	     If (initial ="check") Then
       		Reporter.ReportEvent micFail,"Cell text "&strRowVal&"  should be present in the table","Cell text is not present in the table"
	    End If
    End If
    Environment(strReturnVal1) = intCheck
End Function
'#####################################################################################################
'
'############################################################################################################
'Function name 	    : Func_Wait_Dotnet
'Description        : This function is used for synchronization  with the application
'Parameters       	: The 'Object type' and the 'action being performed is passed as parameters. 
'Assumptions     	: None
'#############################################################################################################
'The following function is used internally.
'#############################################################################################################
Function Func_Wait_DotNet(arrObj,arrKeyValue,initial)
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
'########################################################################################################