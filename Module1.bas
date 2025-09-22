Attribute VB_Name = "Module1"
Sub StartTopLevel()

    

End Sub

Sub RUN_GLOBAL()

    Application.Calculation = xlCalculationManual
    Dim BufferBeforeNext As Variant
    Dim Batch As Integer
    Dim MaxPerDay
    Dim FindNextInvOrShipCounter As Integer
    
    'Get Current batch size
    Sheets("Parameters").Range("K2").Value = ActiveCell.Value
    Sheets("Parameters").Range("K3").Value = ActiveCell.Row
    Sheets("Parameters").Range("K4").Value = ActiveCell.Column
            
    'Loop until end of columns
    Do While ActiveCell.Column < 192
    
        'Color code for Job Number
        If Sheets("Parameters").Range("K7").Value = "Green" Then Sheets("Parameters").Range("K7").Value = "Blue" Else Sheets("Parameters").Range("K7").Value = "Green"
    
        Batch = Sheets("Parameters").Range("K2").Value 'ActiveCell.Value
        'Sheets("Parameters").Range("K2").Value = Batch
        
        'Get Current batch size
        Sheets("Parameters").Range("K2").Value = ActiveCell.Value
        Sheets("Parameters").Range("K3").Value = ActiveCell.Row
        Sheets("Parameters").Range("K4").Value = ActiveCell.Column
        
        'Move to the next row
        If Sheets("Parameters").Range("K7").Value = "Green" Then ActiveCell.Interior.Color = RGB(204, 255, 204) Else ActiveCell.Interior.Color = RGB(153, 204, 255)
        ActiveCell.Offset(1, 0).Activate

    
        'Loop until next empty Cell
        Do While Cells(ActiveCell.Row, 2).Value <> ""
        
            'Gather Buffer before next
            BufferBeforeNext = Cells(ActiveCell.Row, 17).Value
            'Go to the buffeGreen cell
            Do While BufferBeforeNext > 0
                Do While Cells(2, ActiveCell.Column).Value > 5
                    ActiveCell.Offset(0, -1).Activate
                Loop
                ActiveCell.Offset(0, -1).Activate
                Do While Cells(2, ActiveCell.Column).Value > 5
                    ActiveCell.Offset(0, -1).Activate
                Loop
                BufferBeforeNext = BufferBeforeNext - 1
            Loop
    '        ActiveCell.Offset(0, -BufferBeforeNext).Activate
            'Loop untill buffer cell is in a work day
    
            
            'Get Max Per day value
            MaxPerDay = Cells(ActiveCell.Row, 18).Value
            
            Batch = Sheets("Parameters").Range("K2").Value 'ActiveCell.Value
    
            Do While Batch > 0
                'Check if Max per day is less than Batch
                If MaxPerDay < Batch Then
                                                                        If ActiveCell.Value = "" Or ActiveCell.Column <= 20 Then
                                                                            If ActiveCell.Column <= 20 Then
                                                                                GoTo PlanningInThePastNotPossible
                                                                            Else
                                                                                If Sheets("Parameters").Range("K7").Value = "Green" Then ActiveCell.Interior.Color = RGB(204, 255, 204) Else ActiveCell.Interior.Color = RGB(153, 204, 255)
                                                                                ActiveCell.Value = MaxPerDay
                                                                                Call UpdateWCLoad(ActiveCell.Row, ActiveCell.Column - 11, _
                                                                                Cells(ActiveCell.Row, 16).Value, _
                                                                                Cells(ActiveCell.Row, 9).Value, _
                                                                                ActiveCell.Value)
                                                                            End If
                                                                        Else
                                                                            GoTo LogExceptionAlreadyPlanned
                                                                        End If
                    Batch = Batch - MaxPerDay
                    LogCellHistory
                    ActiveCell.Offset(0, -1).Activate
                    Do While MaxPerDay <= Batch
                        Do While Cells(2, ActiveCell.Column).Value > 5
                            ActiveCell.Offset(0, -1).Activate
                        Loop
                                                                        If ActiveCell.Value = "" Or ActiveCell.Column <= 20 Then
                                                                            If ActiveCell.Column <= 20 Then
                                                                                GoTo PlanningInThePastNotPossible
                                                                            Else
                                                                                If Sheets("Parameters").Range("K7").Value = "Green" Then ActiveCell.Interior.Color = RGB(204, 255, 204) Else ActiveCell.Interior.Color = RGB(153, 204, 255)
                                                                                ActiveCell.Value = MaxPerDay
                                                                                Call UpdateWCLoad(ActiveCell.Row, ActiveCell.Column - 11, _
                                                                                Cells(ActiveCell.Row, 16).Value, _
                                                                                Cells(ActiveCell.Row, 9).Value, _
                                                                                ActiveCell.Value)
                                                                            End If
                                                                        Else
                                                                            GoTo LogExceptionAlreadyPlanned
                                                                        End If
                        Batch = Batch - MaxPerDay
                        LogCellHistory
                        If Batch > 0 Then
                            ActiveCell.Offset(0, -1).Activate
                        End If
                    Loop
                    'ActiveCell.Offset(1, 0).Activate
                Else
                    Do While Cells(2, ActiveCell.Column).Value > 5
                        ActiveCell.Offset(0, -1).Activate
                    Loop
                                                                        If ActiveCell.Value = "" Or ActiveCell.Column <= 20 Then
                                                                            If ActiveCell.Column <= 20 Then
                                                                                GoTo PlanningInThePastNotPossible
                                                                            Else
                                                                                If Sheets("Parameters").Range("K7").Value = "Green" Then ActiveCell.Interior.Color = RGB(204, 255, 204) Else ActiveCell.Interior.Color = RGB(153, 204, 255)
                                                                                ActiveCell.Value = Batch
                                                                                Call UpdateWCLoad(ActiveCell.Row, ActiveCell.Column - 11, _
                                                                                Cells(ActiveCell.Row, 16).Value, _
                                                                                Cells(ActiveCell.Row, 9).Value, _
                                                                                ActiveCell.Value)
                                                                            End If
                                                                        Else
                                                                            GoTo LogExceptionAlreadyPlanned
                                                                        End If
                    Batch = Batch - MaxPerDay
                    LogCellHistory
                    'ActiveCell.Offset(1, 0).Activate
                End If
            Loop
            Sheets("Parameters").Range("K8").Value = ActiveCell.Row
            Sheets("Parameters").Range("K9").Value = ActiveCell.Column
            ActiveCell.Offset(1, 0).Activate
        Loop
        
        
        
        
        
        
        
        
        
        
        
        
        
        'Work on sublevels
        If Cells(ActiveCell.Row, 20).Value = "INV" Then
RUNNEXTSUBLEVEL:
            SubLevelsSchedule
            If Cells(ActiveCell.Row, 20).Value = "SHIP DATE" Then GoTo FINISHEDRUNNINGSUBLEVELS
            FindNextInvOrShipCounter = FindNextInvOrShipCounter + 1
            Cells(Sheets("Parameters").Range("K8").Value + 1 + FindNextInvOrShipCounter, Sheets("Parameters").Range("K9").Value).Activate
            'MsgBox Sheets("Parameters").Range("K8").Value + 1 + FindNextInvOrShipCounter
            If Cells(Sheets("Parameters").Range("K8").Value + 1 + FindNextInvOrShipCounter, 20).Value = "SHIP DATE" Then GoTo FINISHEDRUNNINGSUBLEVELS
            Do While Cells(Sheets("Parameters").Range("K8").Value + 1 + FindNextInvOrShipCounter, 20).Value <> "INV" And Cells(Sheets("Parameters").Range("K8").Value + 1 + FindNextInvOrShipCounter, 20).Value <> "SHIP DATE"
                FindNextInvOrShipCounter = FindNextInvOrShipCounter + 1
                Debug.Print Sheets("Parameters").Range("K8").Value + 1 + FindNextInvOrShipCounter
            Loop
            
            If Cells(Sheets("Parameters").Range("K8").Value + 1 + FindNextInvOrShipCounter, 20).Value = "INV" Then
                Cells(Sheets("Parameters").Range("K8").Value + 1 + FindNextInvOrShipCounter, Sheets("Parameters").Range("K9").Value).Activate
                GoTo RUNNEXTSUBLEVEL
            End If
FINISHEDRUNNINGSUBLEVELS:
            If Cells(ActiveCell.Row, 20).Value = "SHIP DATE" Then
                Cells(Sheets("Parameters").Range("K3").Value, Sheets("Parameters").Range("K4").Value).Activate
                FindNextInvOrShipCounter = 0
            End If
            
        Else
            

        
        
        End If
        
        
        
        
        
        
        
        
        
        
        
        
                
GoToNextOrder:
        Cells(Sheets("Parameters").Range("K3").Value, Sheets("Parameters").Range("K4").Value).Activate
        If Range("S1").Value = "Single" Then End
        ActiveCell.Offset(0, 1).Activate

        'Go to next order
        Do While ActiveCell.Value <= 0
            If ActiveCell.Column >= 192 Then GoTo EndOfPlanningWindow
            ActiveCell.Offset(0, 1).Activate
        Loop
    
    Loop
    GoTo NOERROR
    
    
EndOfPlanningWindow:
    End
    
LogExceptionAlreadyPlanned:
    ColorActiveCellRed
    GoTo GoToNextOrder
    
PlanningInThePastNotPossible:
    RunPlanningInThePastNotPossible
    GoTo GoToNextOrder
    
NOERROR:
Application.Calculation = xlCalculationAutomatic
End Sub



Sub SubLevelsSchedule()
    'Get Current batch size
    'Sheets("Parameters").Range("K2").Value = ActiveCell.Value
    Sheets("Parameters").Range("K5").Value = ActiveCell.Row
    Sheets("Parameters").Range("K6").Value = ActiveCell.Column
    
    ActiveCell.Offset(1, 0).Activate
            
    'Loop until end of columns
'    Do While ActiveCell.Column < 192
    
''''''        'Color code for Job Number
''''''        If Sheets("Parameters").Range("K7").Value = "Green" Then Sheets("Parameters").Range("K7").Value = "Blue" Else Sheets("Parameters").Range("K7").Value = "Green"
    
        Batch = Sheets("Parameters").Range("K2").Value 'ActiveCell.Value
        'Sheets("Parameters").Range("K2").Value = Batch
        
        'Get Current batch size
        'Sheets("Parameters").Range("K2").Value = ActiveCell.Value

        
        'Move to the next row
        'If Sheets("Parameters").Range("K7").Value = "Green" Then ActiveCell.Interior.Color = RGB(153,204,0) Else ActiveCell.Interior.Color = RGB(51, 204, 204)
        'ActiveCell.Offset(1, 0).Activate

    
        'Loop until next empty Cell
        Do While Cells(ActiveCell.Row, 2).Value <> ""
        
            'Gather Buffer before next
            BufferBeforeNext = Cells(ActiveCell.Row, 17).Value
            'Go to the buffeGreen cell
            Do While BufferBeforeNext > 0
                Do While Cells(2, ActiveCell.Column).Value > 5
                    ActiveCell.Offset(0, -1).Activate
                Loop
                ActiveCell.Offset(0, -1).Activate
                Do While Cells(2, ActiveCell.Column).Value > 5
                    ActiveCell.Offset(0, -1).Activate
                Loop
                BufferBeforeNext = BufferBeforeNext - 1
            Loop
    '        ActiveCell.Offset(0, -BufferBeforeNext).Activate
            'Loop untill buffer cell is in a work day
            Sheets("Parameters").Range("K5").Value = ActiveCell.Row
            Sheets("Parameters").Range("K6").Value = ActiveCell.Column
            
            'Get Max Per day value
            MaxPerDay = Cells(ActiveCell.Row, 18).Value
            
            Batch = Sheets("Parameters").Range("K2").Value 'ActiveCell.Value
    
            Do While Batch > 0
                'Check if Max per day is less than Batch
                If MaxPerDay < Batch Then
                                                                        If ActiveCell.Value = "" Or ActiveCell.Column <= 20 Then
                                                                            If ActiveCell.Column <= 20 Then
                                                                                GoTo PlanningInThePastNotPossible
                                                                            Else
                                                                                If Sheets("Parameters").Range("K7").Value = "Green" Then ActiveCell.Interior.Color = RGB(153, 204, 0) Else ActiveCell.Interior.Color = RGB(51, 204, 204)
                                                                                ActiveCell.Value = MaxPerDay
                                                                                Call UpdateWCLoad(ActiveCell.Row, ActiveCell.Column - 11, _
                                                                                Cells(ActiveCell.Row, 16).Value, _
                                                                                Cells(ActiveCell.Row, 9).Value, _
                                                                                ActiveCell.Value)
                                                                            End If
                                                                        Else
                                                                            GoTo LogExceptionAlreadyPlanned
                                                                        End If
                    Batch = Batch - MaxPerDay
                    LogCellHistory
                    ActiveCell.Offset(0, -1).Activate
                    Do While MaxPerDay <= Batch
                        Do While Cells(2, ActiveCell.Column).Value > 5
                            ActiveCell.Offset(0, -1).Activate
                        Loop
                                                                        If ActiveCell.Value = "" Or ActiveCell.Column <= 20 Then
                                                                            If ActiveCell.Column <= 20 Then
                                                                                GoTo PlanningInThePastNotPossible
                                                                            Else
                                                                                If Sheets("Parameters").Range("K7").Value = "Green" Then ActiveCell.Interior.Color = RGB(153, 204, 0) Else ActiveCell.Interior.Color = RGB(51, 204, 204)
                                                                                ActiveCell.Value = MaxPerDay
                                                                                Call UpdateWCLoad(ActiveCell.Row, ActiveCell.Column - 11, _
                                                                                Cells(ActiveCell.Row, 16).Value, _
                                                                                Cells(ActiveCell.Row, 9).Value, _
                                                                                ActiveCell.Value)
                                                                            End If
                                                                        Else
                                                                            GoTo LogExceptionAlreadyPlanned
                                                                        End If
                        Batch = Batch - MaxPerDay
                        LogCellHistory
                        If Batch > 0 Then
                            ActiveCell.Offset(0, -1).Activate
                        End If
                    Loop
                    'ActiveCell.Offset(1, 0).Activate
                Else
                    Do While Cells(2, ActiveCell.Column).Value > 5
                        ActiveCell.Offset(0, -1).Activate
                    Loop
                                                                        If ActiveCell.Value = "" Or ActiveCell.Column <= 20 Then
                                                                            If ActiveCell.Column <= 20 Then
                                                                                GoTo PlanningInThePastNotPossible
                                                                            Else
                                                                                If Sheets("Parameters").Range("K7").Value = "Green" Then ActiveCell.Interior.Color = RGB(153, 204, 0) Else ActiveCell.Interior.Color = RGB(51, 204, 204)
                                                                                ActiveCell.Value = Batch
                                                                                Call UpdateWCLoad(ActiveCell.Row, ActiveCell.Column - 11, _
                                                                                Cells(ActiveCell.Row, 16).Value, _
                                                                                Cells(ActiveCell.Row, 9).Value, _
                                                                                ActiveCell.Value)
                                                                            End If
                                                                        Else
                                                                            GoTo LogExceptionAlreadyPlanned
                                                                        End If
                    Batch = Batch - MaxPerDay
                    LogCellHistory
                    'ActiveCell.Offset(1, 0).Activate
                End If
            Loop
            ActiveCell.Offset(1, 0).Activate
        Loop
    'Loop
    GoTo NOERROR
    
GoToNextOrder:
'''        Cells(Sheets("Parameters").Range("K3").Value, Sheets("Parameters").Range("K4").Value).Activate
'''        If Range("S1").Value = "Single" Then End
'''        ActiveCell.Offset(0, 1).Activate
'''
'''        'Go to next order
'''        Do While ActiveCell.Value <= 0
'''            If ActiveCell.Column >= 192 Then GoTo EndOfPlanningWindow
'''            ActiveCell.Offset(0, 1).Activate
'''        Loop
'''
'''    'Loop
    GoTo NOERROR
    
    
EndOfPlanningWindow:
    End
    
LogExceptionAlreadyPlanned:
    ColorActiveCellRed
    GoTo GoToNextOrder
    
PlanningInThePastNotPossible:
    RunPlanningInThePastNotPossible
    GoTo GoToNextOrder
    
NOERROR:

    
End Sub




Sub RunPlanningInThePastNotPossible()
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Interior.Color = RGB(255, 0, 0) ' Red color
    Sheets("Log").Range("A" & Sheets("Parameters").Range("A2").Value).Value = ActiveCell.Row
    Sheets("Log").Range("B" & Sheets("Parameters").Range("A2").Value).Value = ActiveCell.Column
    Sheets("Log").Range("C" & Sheets("Parameters").Range("A2").Value).Value = Sheets("Parameters").Range("K3").Value
    Sheets("Log").Range("D" & Sheets("Parameters").Range("A2").Value).Value = Sheets("Parameters").Range("K4").Value
    Sheets("Log").Range("E" & Sheets("Parameters").Range("A2").Value).Value = "Error Trying to Plan in the Past"
    Sheets("Parameters").Range("A2").Value = Sheets("Parameters").Range("A2").Value + 1
    Cells(Sheets("Parameters").Range("K3").Value, Sheets("Parameters").Range("K4").Value).Activate
    ActiveCell.Interior.Color = RGB(255, 0, 0) ' Red color

End Sub

Sub ColorActiveCellRed()
    ActiveCell.Interior.Color = RGB(255, 0, 0) ' Red color
    Sheets("Log").Range("A" & Sheets("Parameters").Range("A2").Value).Value = ActiveCell.Row
    Sheets("Log").Range("B" & Sheets("Parameters").Range("A2").Value).Value = ActiveCell.Column
    Sheets("Log").Range("C" & Sheets("Parameters").Range("A2").Value).Value = Sheets("Parameters").Range("K3").Value
    Sheets("Log").Range("D" & Sheets("Parameters").Range("A2").Value).Value = Sheets("Parameters").Range("K4").Value
    Sheets("Log").Range("E" & Sheets("Parameters").Range("A2").Value).Value = "Over Capacity Plan Stop - Review"
    Sheets("Parameters").Range("A2").Value = Sheets("Parameters").Range("A2").Value + 1
    Cells(Sheets("Parameters").Range("K3").Value, Sheets("Parameters").Range("K4").Value).Activate
    ActiveCell.Interior.Color = RGB(255, 0, 0) ' Red color

End Sub

Sub LogCellHistory()

    Sheets("Log").Range("A" & Sheets("Parameters").Range("A2").Value).Value = ActiveCell.Row
    Sheets("Log").Range("B" & Sheets("Parameters").Range("A2").Value).Value = ActiveCell.Column
    Sheets("Log").Range("C" & Sheets("Parameters").Range("A2").Value).Value = Sheets("Parameters").Range("K3").Value
    Sheets("Log").Range("D" & Sheets("Parameters").Range("A2").Value).Value = Sheets("Parameters").Range("K4").Value
    Sheets("Log").Range("F" & Sheets("Parameters").Range("A2").Value).Value = Sheets("Parameters").Range("K5").Value
    Sheets("Log").Range("G" & Sheets("Parameters").Range("A2").Value).Value = Sheets("Parameters").Range("K6").Value
    Sheets("Log").Range("E" & Sheets("Parameters").Range("A2").Value).Value = "Production Planned"
    Sheets("Parameters").Range("A2").Value = Sheets("Parameters").Range("A2").Value + 1
End Sub

Sub GoToNextDelivery()
    
    
    
End Sub



Sub UpdateWCLoad(targetRow As Long, targetCol As Long, loadRowRef As Variant, qty As Variant, multiplier As Variant)
    Dim loadRow As Long
    Dim val1 As Double, val2 As Double, val3 As Double
    
    ' Validate row reference from Column 16
    If IsNumeric(loadRowRef) And loadRowRef > 0 Then
        loadRow = CLng(loadRowRef)
    Else
        MsgBox "Invalid row reference in column 16.", vbExclamation
        Exit Sub
    End If
    
    ' Get existing value from WC Load
    If Not IsError(Sheets("WC Load").Cells(loadRow, targetCol).Value) And IsNumeric(Sheets("WC Load").Cells(loadRow, targetCol).Value) Then
        val1 = CDbl(Sheets("WC Load").Cells(loadRow, targetCol).Value)
    End If
    
    ' Get quantity (Column 9 value)
    If Not IsError(qty) And IsNumeric(qty) Then
        val2 = CDbl(qty)
    End If
    
    ' Get multiplier (ActiveCell value)
    If Not IsError(multiplier) And IsNumeric(multiplier) Then
        val3 = CDbl(multiplier)
    End If
    
    ' Perform calculation and update WC Load
    Sheets("WC Load").Cells(loadRow, targetCol).Value = val1 + (val2 * val3)
End Sub

