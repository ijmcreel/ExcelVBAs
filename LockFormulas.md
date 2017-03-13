		Sub ProtectFormulas()

				'ctrl+shift+L


		 On Error GoTo Handler:


		 Application.ScreenUpdating = False

			If ActiveSheet.Unprotect Then
			  ActiveSheet.Unprotect
			Else
				Cells.Select
			        Selection.Locked = False
				Selection.SpecialCells(xlCellTypeFormulas, 23).Select
				Selection.Locked = True
				ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
				Range("A1").Select
		          End If

		          Exit Sub
				Handler:    MsgBox ("there are no formulas here, what the hell do you think you are doing?!")
				Range("a1").Select
		          End Sub
