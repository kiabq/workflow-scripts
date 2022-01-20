Attribute VB_Name = "AutoLabel"
' TODO _
    - Loop through open worksheets and match them with regex or instr _
    - Add Refresh date macro to automatically refresh a date (preferably with a shortcut) in a certain range of cell(s). _

' Also, Option Explicit should really be enabled here, just make sure to do that and define all veriables future me. (if you havent already)

Option Explicit

'Fix for bug involving the dataObject in VBA and it's issue with the .PutInClipboard method.
' https://wellsr.com/vba/2015/tutorials/vba-copy-to-clipboard-paste-clear/

#If Mac Then
    ' do nothing
#Else
    #If VBA7 Then
        Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                                             ByVal dwBytes As LongPtr) As LongPtr

        Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
        Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
        Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long

        Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
                                                         ByVal lpString2 As Any) As LongPtr

        Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat _
                                                                As Long, ByVal hMem As LongPtr) As LongPtr
    #Else
        Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                                     ByVal dwBytes As Long) As Long

        Declare Function CloseClipboard Lib "User32" () As Long
        Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
        Declare Function EmptyClipboard Lib "User32" () As Long

        Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
                                                 ByVal lpString2 As Any) As Long

        Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
                                                        As Long, ByVal hMem As Long) As Long
    #End If
#End If
Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Sub ClipBoard_SetData(MyString As String)
    #If Mac Then
        With New MSForms.DataObject
            .SetText MyString
            .PutInClipboard
        End With
    #Else
        #If VBA7 Then
            Dim hGlobalMemory As LongPtr
            Dim hClipMemory   As LongPtr
            Dim lpGlobalMemory    As LongPtr
        #Else
            Dim hGlobalMemory As Long
            Dim hClipMemory   As Long
            Dim lpGlobalMemory    As Long
        #End If

        Dim x                 As Long

        ' Allocate moveable global memory.
       '-------------------------------------------
       hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

        ' Lock the block to get a far pointer
       ' to this memory.
       lpGlobalMemory = GlobalLock(hGlobalMemory)

        ' Copy the string to this global memory.
       lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

        ' Unlock the memory.
       If GlobalUnlock(hGlobalMemory) <> 0 Then
            MsgBox "Could not unlock memory location. Copy aborted."
            GoTo PrepareToClose
        End If

        ' Open the Clipboard to copy data to.
       If OpenClipboard(0&) = 0 Then
            MsgBox "Could not open the Clipboard. Copy aborted."
            Exit Sub
        End If

        ' Clear the Clipboard.
       x = EmptyClipboard()

        ' Copy the data to the Clipboard.
       hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

PrepareToClose:

        If CloseClipboard() = 0 Then
            MsgBox "Could not close Clipboard."
        End If
    #End If

End Sub

Function reportType()

    Dim reportName As String
    Dim reportNum As Integer
    reportName = ActiveSheet.Name
    reportNum = InStr(1, reportName, "Original")
    If reportNum > 0 Then reportType = "80" Else reportType = "84"

End Function

 

Private Sub Obtain()

    Dim reportType As String, reportLocal As Integer, reportNum As Integer, reportName As String, reportRow
    reportName = ActiveSheet.Name
    reportLocal = InStr(1, reportName, "Local")
    reportNum = InStr(1, reportName, "Original")
    
    ' Define values (reportType) to append to strings in your defined rows (reportRow)
    If reportLocal > 0 Then

        If reportNum > 0 Then

            reportRow = "A"

            reportType = "80"

        Else

            reportRow = "A"

            reportType = "84"

        End If

    End If

    If reportLocal = 0 Then

        If reportNum > 0 Then

            reportRow = "A"

            reportType = "80"

        Else

            reportRow = "A"

            reportType = "84"

        End If

    End If

    Dim v As String
    Dim i As Integer, j As Integer, counter As Integer, size As Integer

    Dim selectedCells()
    j = Range("A2").Row
    i = Range("A" & Rows.Count).End(xlUp).Row
    size = (i - j)
    
    ReDim selectedCells(size)
    v = ""

    For counter = j To i
        selectedCells(counter - j) = "20" + CStr(Range("A" & counter)) + reportType
    Next counter

    v = Join(selectedCells, ",")
    Debug.Print (v)
    ClipBoard_SetData (v)

End Sub

Sub Build()

    Call Obtain

End Sub


