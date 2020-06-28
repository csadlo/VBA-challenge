Attribute VB_Name = "Module1"
Sub Process_Stock_Data()

    ' Set an initial variable for holding the ticker name
    Dim Ticker_Code As String
    Dim Yearly_Change As Double
    Dim Percent_Chance As Double
    Dim Volume_Traded As Double
    
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume As Double
    Dim Greatest_Percent_Increase_Ticker_Code As String
    Dim Greatest_Percent_Decrease_Ticker_Code As String
    Dim Greatest_Total_Volume_Ticker_Code As String
    
    Dim First_Day_OTY_Opening_Value As Double
    Dim Final_Day_OTY_Closing_Value As Double
    

    ' Book-keeping variables
    Dim Ticker_Code_Entry_Count As Long
    Dim Unique_Ticker_Code_Count As Long
    Dim Summary_Table_Row As Long
    First_Time_Through = False

    Yearly_Change = 0
    Percent_Chance = 0
    Volume_Traded = 0
    Greatest_Percent_Increase = 0
    Greatest_Percent_Decrease = 0
    Greatest_Total_Volume = 0
    
    First_Day_OTY_Opening_Value = 0
    Final_Day_OTY_Closing_Value = 0

    Unique_Ticker_Code_Count = 0
    Summary_Table_Row = 1

    ' Create all of the Summary Table Labels
    Range("K" & Summary_Table_Row).Value = "Ticker"
    Range("L" & Summary_Table_Row).Value = "Yearly Change"
    Range("M" & Summary_Table_Row).Value = "Percent Change"
    Range("N" & Summary_Table_Row).Value = "Total Stock Volume"
    Range("Q" & Summary_Table_Row + 1).Value = "Greatest Percent Increase"
    Range("Q" & Summary_Table_Row + 2).Value = "Greatest Percent Decrease"
    Range("Q" & Summary_Table_Row + 3).Value = "Greatest Total Volume"
    Range("R" & Summary_Table_Row).Value = "Ticker"
    Range("S" & Summary_Table_Row).Value = "Value"

    ' Debug.Print ("Begin Processing")


    Set sh = ActiveSheet
    
    ' Loop through all stock movement data entries
    'For rw = 1 To 5000
    For Each rww In sh.Rows
       rw = rww.Row

        'Debug.Print ("First Ticker Code is " + Cells(rw, 1).Value)
        'Debug.Print ("Second Ticker Code is " + Cells(rw - 1, 1).Value)

        If rw <= 1 Then
            ' A judicious continue since VBA doesnt support Continue
            GoTo NextIteration
        End If

        ' Check if we are still within the same stock ticker name, if not, then...
        ' If we just discovered a new ticker code...
        If Cells(rw, 1).Value <> Cells(rw - 1, 1).Value Then
            
            ' Set the Ticker name
            Ticker_Code = Cells(rw, 1).Value
        
            ' Print the Ticker Code in the Summary Table
            cellrange = "K" & Summary_Table_Row + 1
            Range(cellrange).Value = Ticker_Code
            
            
            ' We want to skip this bit of code if we are comparing with the header cell as the previous cell
            ' This code is basically starts running once we reach the end of the first ticker's set of entries
            ' and have started peeking at the next Ticker Code's set of entries
            If Header_Finished Then
                
                ' Track the (old ticker code)
                Final_Day_OTY_Closing_Value = Cells(rw - 1, 6)
                
                ' MsgBox "Closing value for " & Ticker_Code & " is " & Final_Day_OTY_Closing_Value
                
                Yearly_Change = Final_Day_OTY_Closing_Value - First_Day_OTY_Opening_Value
                Percent_Change = (Yearly_Change / First_Day_OTY_Opening_Value) * 100
                
                ' Update the Ticker Yearly Change in the Summary Table
                cellrange = "L" & Summary_Table_Row
                Range(cellrange).Value = Yearly_Change
                Range(cellrange).NumberFormat = "#,##0.00"
                
                ' Do conditional formatting here
                If Yearly_Change > 0 Then
                    Range(cellrange).Interior.ColorIndex = 4    ' Green
                ElseIf Yearly_Change < 0 Then
                    Range(cellrange).Interior.ColorIndex = 3    ' Red
                End If

                
                ' Update the Ticker Percent Change in the Summary Table
                cellrange = "M" & Summary_Table_Row
                Range(cellrange).Value = Percent_Change
                Range(cellrange).NumberFormat = "#,##0.00"
                
                ' Update the Ticker Volume Traded in the Summary Table
                cellrange = "N" & Summary_Table_Row
                Range(cellrange).Value = Volume_Traded
                
                ' Update if better was found
                If Percent_Change > Greatest_Percent_Increase Then
                    Greatest_Percent_Increase = Percent_Change
                    Greatest_Percent_Increase_Ticker_Code = Cells(rw - 1, 1)
                    'MsgBox "Greatest Percent Increase is " & Greatest_Percent_Increase & " from " & Greatest_Percent_Increase_Ticker_Code
                    ' Update Greatest Display Table
                    Range("R" & 2).Value = Greatest_Percent_Increase_Ticker_Code
                    Range("S" & 2).Value = Greatest_Percent_Increase
                End If
                
                ' Update if better was found
                If Percent_Change < Greatest_Percent_Decrease Then
                    Greatest_Percent_Decrease_Ticker_Code = Cells(rw - 1, 1)
                    Greatest_Percent_Decrease = Percent_Change
                    'MsgBox "Greatest Percent Decrease is " & Greatest_Percent_Decrease & " from " & Greatest_Percent_Decrease_Ticker_Code
                    ' Update Greatest Display Table
                    Range("R" & 3).Value = Greatest_Percent_Decrease_Ticker_Code
                    Range("S" & 3).Value = Greatest_Percent_Decrease
                End If
                
                ' Update if better was found
                If Volume_Traded > Greatest_Total_Volume Then
                    Greatest_Total_Volume_Ticker_Code = Cells(rw - 1, 1)
                    Greatest_Total_Volume = Volume_Traded
                    'MsgBox "Greatest Total Volume is " & Greatest_Total_Volume & " from " & Greatest_Total_Volume_Ticker_Code
                    ' Update Greatest Display Table
                    Range("R" & 4).Value = Greatest_Total_Volume_Ticker_Code
                    Range("S" & 4).Value = Greatest_Total_Volume
                End If
    
            End If
            Header_Finished = True
            
            ' Track the (new ticker code)
            First_Day_OTY_Opening_Value = Cells(rw, 3)

            ' Increment
            Unique_Ticker_Code_Count = Unique_Ticker_Code_Count + 1
            Summary_Table_Row = Summary_Table_Row + 1
            ' Debug.Print (Summary_Table_Row)

            ' Reset Values
            Ticker_Code_Entry_Count = 0
            Volume_Traded = 0
                
            ' Add to the Volume_Traded
            Volume_Traded = Volume_Traded + Cells(rw, 7).Value

            ' If the cell immediately following a row is the same Ticker Code...
        End If

        ' Track how many entries we have for this specific Ticker Code
        Ticker_Code_Entry_Count = Ticker_Code_Entry_Count + 1
            
        ' Add to the Volume_Traded
        Volume_Traded = Volume_Traded + Cells(rw, 7).Value


NextIteration:

    Next

End Sub

Sub Reset_Summary_Table()

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 1
    
    Set sh = ActiveSheet
    
    sh.Columns(11).ClearContents
    sh.Columns(12).ClearContents
    sh.Columns(12).FormatConditions.Delete
    sh.Columns(13).ClearContents
    sh.Columns(14).ClearContents
    sh.Columns(18).ClearContents
    sh.Columns(19).ClearContents
    
    Range("K" & Summary_Table_Row).Value = "Ticker"
    Range("L" & Summary_Table_Row).Value = "Yearly Change"
    Range("M" & Summary_Table_Row).Value = "Percent Change"
    Range("N" & Summary_Table_Row).Value = "Total Stock Volume"
    
    
    'For Each rww In sh.Rows
    '    rw = rww.Row
    
    '    Range("K" & rw).ClearContents
    '    Range("L").ClearContents
    
    'Next
    
End Sub
