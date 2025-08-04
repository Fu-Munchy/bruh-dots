REM  *****  BASIC  *****


'***** TODO *****
' - [X] Find a way to perform calculations, and paste result into a given cell.
' - [] Maybe incorporate the TimesReviewed variable into the algorithim?
' - [X] Find a way to execute the Macro by clicking a button.


Function CalcStudyDate()
'	*** DATES ***
	Dim LastStudied As Object ' Date of the most recent study session - found in Col C
	Dim NextStudyDate As Variant	 ' When I should study this topic next
	Dim DeadlineForStudy As Object
	
'	*** Score and Algorithm variables ***
	Dim LastTestScore As Object ' The score (in perecentage) of the most recent test - found in Col E
	Dim TimesReviewed As Object ' Number of times reviewed the topic - found in the 5th Column - Col F
	Dim TopicDifficulty As Single ' Intervals for various scores

'	*** Cell and Document Manipulation variables ***
	Dim CurrDocument As Object ' The current active Document
	Dim CurrSheet As Object ' The current Sheet within the Document
	Dim CurrCellSelection As Object ' Selection within the Current Document and Sheet.
	Dim TodayDate As Object ' Today's date as a reference 
	

'	************************************************
'	*** Processing input of cells into variables ***
'	************************************************
	
	' The CurrDocument is the currently open LibreOffice Document
	CurrDocument = ThisComponent
	
	' CurrSheet is the ACTIVE sheet within the aforementioned Document.
	CurrSheet = CurrDocument.CurrentController.ActiveSheet
		
	' The origin Cell, where counting starts
	OriginCell = CurrSheet.GetCellRangeByName("A1")
	
	' Today's date location
	TodayDate = CurrSheet.GetCellRangeByName("L2")
	
	
	'The last used row in the current sheet in the current document
	LastUsedRow = CurrSheet.createCursorByRange(OriginCell)
	LastUsedRow.gotoEnd()
	
	' For each row in the Sheet, check if the cell is empty. True => Proceed to apply variables from other columns. False => Error Message.
	For i=1 To LastUsedRow.RangeAddress.EndRow Step 1
	
		CurrCellSelection = CurrSheet.getCellByPosition(6,i)
		LastStudied = CurrSheet.getCellByPosition(2,i)
		LastTestScore = CurrSheet.getCellByPosition(4,i)
		TimesReviewed = CurrSheet.getCellByPosition(5,i)
		DeadlineForStudy = CurrSheet.getCellByPosition(7,i)
		
		Select Case CurrCellSelection.Value
			Case Is = 0
				'	*******************************************
				'	*** Algorithm for Calculating next date ***
				'	*******************************************	
				' Schedule the next study date based on the score
				Select Case LastTestScore.Value
					' If the score is between 70-80%, then NextStudyDate is in 1-2 weeks
					Case 0.60 To 0.80
						' A random number between [6,14]
						TopicDifficulty = Int((Rnd()*7) + 6) 
						NextStudyDate = TopicDifficulty + LastStudied.Value
						CurrCellSelection.Value = NextStudyDate
						
						' If the Score is between 60-80% but the topic has only been reviewed once or less, then schedule the time for tomorrow.
						if TimesReviewed.Value <= 1 Then	
							TopicDifficulty = 1 
							NextStudyDate = TopicDifficulty + LastStudied.Value
							CurrCellSelection.Value = NextStudyDate
						End If
						
					' If the score is between 80-90% then NextStudyDate is in 2-4 weeks	
					Case 0.81 To 0.85
						' A random number between [16,30]
						TopicDifficulty = Int((Rnd()*16) + 15) 
						NextStudyDate = TopicDifficulty + LastStudied.Value
						CurrCellSelection.Value = NextStudyDate
						
							' If the Score is between 60-80% but the topic has only been reviewed once or less, then schedule the time for tomorrow.
						if TimesReviewed.Value <= 1 Then	
							TopicDifficulty = 1 
							NextStudyDate = TopicDifficulty + LastStudied.Value
							CurrCellSelection.Value = NextStudyDate
						End If
						
					' If score is 90% or above, then NextStudyDate is in 60-100 days
					Case 0.86 To 1.00  
						' A random number between [60,99]
						TopicDifficulty = Int((Rnd()*40) + 60) 
						NextStudyDate = TopicDifficulty + LastStudied.Value
						CurrCellSelection.Value = NextStudyDate
						'Print TopicDifficulty
						
							' If the Score is between 60-80% but the topic has only been reviewed once or less, then schedule the time for tomorrow.
						if TimesReviewed.Value <= 1 Then	
							TopicDifficulty = 1 
							NextStudyDate = TopicDifficulty + LastStudied.Value
							CurrCellSelection.Value = NextStudyDate
						End If
						
					Case 0.00 To 0.59
						' 1 Day until next revision
						TopicDifficulty = 1 
						NextStudyDate = TopicDifficulty + LastStudied.Value
						CurrCellSelection.Value = NextStudyDate							
				End Select
				
				' If the NextStudyDate is later than the DeadlineForStudy date, then schedule the NextStudy Date before the deadline.
				If NextStudyDate > DeadlineForStudy.Value Then
					CurrCellSelection.Value = NextStudyDate - (NextStudyDate - DeadlineForStudy.Value) + 1

				End If 
				
			' Check if the CurrCellSelection display's today's date => The review deadline is today, and needs to be updated.
			Case TodayDate.Value
					LastStudied.CellBackColor = RGB(128,0,0) 
					CurrCellSelection.CellBackColor = RGB(128,0,0) 
					Print "Update the marked dates!"
		End Select
	Next
	
	

End Function


