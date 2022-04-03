'Datatable.AddSheet "Module"
'Datatable.ImportSheet "C:\Users\sfjbs\Desktop\UFT\Framework\KeywordDrivenFramework\Organizer\Organizer.xlsx",1,"Module"

Services.StartTransaction "tr1"


mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount

For i = 1  To mrowcount Step 1
	Datatable.SetCurrentRow(i)
	
	Modexe =  Datatable("ModuleExe","Action1")
	
	'msgbox Modexe
	If Modexe="Y" Then
		Modid=Datatable("ModuleID","Action1")
		msgbox Modid
		
		trowcount=datatable.GetSheet("Action2").GetRowCount
		
		msgbox trowcount
		
		For j = 1 To trowcount Step 1
			Datatable.SetCurrentRow(j)
			If Modid=Datatable("ModuleID","Action2") and Datatable("TestCaseExe","Action2") = "Y"  Then
			testcaseid=Datatable("TestCaseID","Action2")	
			msgbox testcaseid
			
			tsrowcount=Datatable.GetSheet("Action3").GetRowCount
			msgbox tsrowcount
		For k = 1 To tsrowcount Step 1
			datatable.SetCurrentRow(k)
			If testcaseid=Datatable("TestCaseID","Action3") Then
			keyword=Datatable("Keyword","Action3")
			msgbox keyword
			
			Select Case (keyword)
				
				Case "In"
				Call login("john","hp")
				
				Case "ca"
				Call closeapp()
				
				Case "oo"
				Call openOrder("5")
				
				Case "uo"
				Call updateOrder()
				
				Case "lnd"
			        drowcount = datatable.GetSheet("Action4").GetRowCount
			        
			        For l = 1 To drowcount Step 1
			        	datatable.SetCurrentRow(l)
			        	Call login(datatable("username","Action4"),datatable("password","Action4"))
			        	Call closeapp()
			        Next
			        
			        case "ood"
			        orrowcount = datatable.GetSheet("Action4").GetRowCount
			        
			        For m = 1 To orrowcount Step 1
			        	datatable.SetCurrentRow(m)
			        	Call openOrder(datatable("orderno","Action4"))
			        	
			        Next
				
			End Select				
			End If
		Next
			End If
		Next
	End If
Next



 @@ hightlight id_;_1890032936_;_script infofile_;_ZIP::ssf7.xml_;_

Services.EndTransaction "tr1"

 @@ hightlight id_;_2119708240_;_script infofile_;_ZIP::ssf8.xml_;_



