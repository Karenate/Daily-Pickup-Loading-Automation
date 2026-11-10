Attribute VB_Name = "PickupImport"
Option Explicit

' PUBLIC PORTFOLIO DEMO — synthetic workbook layout and data only.
' Expected sheets: Pickup_A, Pickup_B, Pickup_C, Daily_A, Daily_B, Daily_C.
' Source: A=Shipment ID, B=Delivered Date, C=Net Weight (LB).
' Destination: A=Daily Date/Time, B=Truck Count, C=Shipment IDs, D=Total Net Weight.
' Dates are matched by calendar day, ignoring time. Only matched destination dates are written.

Public Sub Import_All()
    Dim msg As String
    msg = "Synthetic Pickup Import Summary" & vbCrLf & vbCrLf
    msg = msg & ImportProduct("A") & vbCrLf
    msg = msg & ImportProduct("B") & vbCrLf
    msg = msg & ImportProduct("C")
    MsgBox msg, vbInformation, "Import Summary"
End Sub

Public Sub Import_A()
    MsgBox ImportProduct("A"), vbInformation, "Product A Summary"
End Sub
Public Sub Import_B()
    MsgBox ImportProduct("B"), vbInformation, "Product B Summary"
End Sub
Public Sub Import_C()
    MsgBox ImportProduct("C"), vbInformation, "Product C Summary"
End Sub

Private Function ImportProduct(ByVal productKey As String) As String
    Dim sourceSheet As Worksheet, targetSheet As Worksheet
    Dim counts As Object, weights As Object, ids As Object
    Dim lastSource As Long, lastTarget As Long, r As Long
    Dim k As String, shipmentId As String, wt As Double
    Dim daysUpdated As Long, trucksUpdated As Long, weightUpdated As Double
    Dim priorEvents As Boolean, priorScreen As Boolean
    Dim priorCalculation As XlCalculation
    Dim completed As Boolean, errorMessage As String

    priorEvents = Application.EnableEvents
    priorScreen = Application.ScreenUpdating
    priorCalculation = Application.Calculation
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set sourceSheet = ThisWorkbook.Worksheets("Pickup_" & productKey)
    Set targetSheet = ThisWorkbook.Worksheets("Daily_" & productKey)
    Set counts = CreateObject("Scripting.Dictionary")
    Set weights = CreateObject("Scripting.Dictionary")
    Set ids = CreateObject("Scripting.Dictionary")

    lastSource = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).Row
    For r = 2 To lastSource
        k = DayKey(sourceSheet.Cells(r, "B").Value)
        shipmentId = Trim$(CStr(sourceSheet.Cells(r, "A").Value))
        If Len(k) > 0 And Len(shipmentId) > 0 Then
            If Not IsNumeric(sourceSheet.Cells(r, "C").Value) Then Err.Raise vbObjectError + 101, , "Invalid weight at source row " & r
            wt = CDbl(sourceSheet.Cells(r, "C").Value)
            If wt < 0 Then Err.Raise vbObjectError + 102, , "Negative weight at source row " & r
            If Not counts.Exists(k) Then
                counts.Add k, 0&
                weights.Add k, 0#
                ids.Add k, ""
            End If
            counts(k) = CLng(counts(k)) + 1
            weights(k) = CDbl(weights(k)) + wt
            If Len(CStr(ids(k))) > 0 Then ids(k) = ids(k) & ", "
            ids(k) = ids(k) & shipmentId
        End If
    Next r

    lastTarget = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row
    If lastTarget >= 2 Then targetSheet.Range("B2:D" & lastTarget).ClearContents
    For r = 2 To lastTarget
        k = DayKey(targetSheet.Cells(r, "A").Value)
        If Len(k) > 0 And counts.Exists(k) Then
            targetSheet.Cells(r, "B").Value = counts(k)
            targetSheet.Cells(r, "C").Value = ids(k)
            targetSheet.Cells(r, "D").Value = weights(k)
            daysUpdated = daysUpdated + 1
            trucksUpdated = trucksUpdated + CLng(counts(k))
            weightUpdated = weightUpdated + CDbl(weights(k))
        End If
    Next r
    If lastTarget >= 2 Then targetSheet.Range("D2:D" & lastTarget).NumberFormat = "#,##0"
    completed = True
Cleanup:
    Application.Calculation = priorCalculation
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorScreen
    If completed Then
        ImportProduct = "Product " & productKey & ": " & daysUpdated & " days | " & trucksUpdated & " trucks | " & Format$(weightUpdated, "#,##0") & " LB"
    Else
        ImportProduct = "Product " & productKey & ": ERROR - " & errorMessage
    End If
    Exit Function
Failed:
    errorMessage = Err.Description
    Resume Cleanup
End Function

Private Function DayKey(ByVal v As Variant) As String
    DayKey = ""
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    If IsDate(v) Then DayKey = Format$(CDate(v), "yyyy-mm-dd")
End Function
