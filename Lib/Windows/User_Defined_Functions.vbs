'###############################  KEYWORD DRIVEN FRAMEWORK   #####################################
'File Name			: User Defined Functions 
'Author		       	: Open2Test
'Version	    	: V 2.1
'Date of Creation	: 12-Mar-2013
'######################################  Common Functions  #######################################
Option Explicit 	'To enforce variable declaration
'#################################################################################################
'Function name 		: Func_FunctionCall
'Description        : If the user requires to perform application specific operations /functions.
'Parameters       	: 1. The Function Name to be used.
'				  	  2. The parameters to be used with the called Function.                        
'Assumptions     	: NA
'#################################################################################################
'The following function is for 'CallFunction' Keyword
'#################################################################################################
Public Function Func_FunctionCall()
   ReDim arrActionParam(1)	  	'Stores the Action Parameters  
   Dim intActionParamCount	'Stores the number of parameters 
   Dim inta					'Used for looping
   Dim strFunName
	Dim i
	Dim ParCnt
	If (DataTable.Value(4,dtLocalSheet)<> "") then
   ParCnt = Ubound(arrKeyIndex)
	ReDim arrActionParam(ParCnt)
	For i = 0 to ParCnt
		If Lcase(arrKeyIndex(i)) = "null" Then
			arrActionParam(i) = ""
		Else
        arrActionParam(i) = cstr(arrKeyIndex(i))
		End if
	Next
 End If
 
  Select Case lcase(arrAction(0)) 'Selecting the used defined functions
   Case "function1"
		Func_FunctionCall = func_Example1(arrActionParam(0))
   Case "function2"
		call func_Function2()
	Case else
		Reporter.ReportEvent micFail, "User Defined Function mentioned in the row # " & intRowCount & " does not exist", "Please check the keyword."
  End select  
End function
'#################################################################################################
Function func_Example1(test1)
	'Write the User Defined function here
	Msgbox test1
	func_Example1 = 1
End Function
'#################################################################################################
Function func_Function2()
	'Write the User Defined function here
	Msgbox "Open2Test"
End Function
'#################################################################################################