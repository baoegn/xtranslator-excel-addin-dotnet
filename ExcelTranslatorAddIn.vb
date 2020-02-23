Imports Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.ComponentModel

Public Class ExcelTranslatorAddIn

    ' Handle worker threads
    Private bgHandler As BackgroundWorker = New BackgroundWorker

    ' Translate class
    Private WithEvents translator1 As cTranslator
    Private WithEvents translator2 As cTranslator
    Private WithEvents translator3 As cTranslator
    Private WithEvents translator4 As cTranslator

    ' Identify busy or not
    Private isBusy1 As Boolean
    Private isBusy2 As Boolean
    Private isBusy3 As Boolean
    Private isBusy4 As Boolean

    Private workingIndexes As New List(Of String) ' Store working cell index
    Private curWorkingIndex As Integer

    Private curCells As cellProperties
    Private undoCells As New List(Of cellProperties) ' For undo transaction
    Private Const MAX_UNDO = 10 ' Ten times undo

    Private lngFrom As String
    Private lngTo As String
    Private lngCharSet As String

    Private Sub setLabel(ByVal show As Boolean, ByVal str As String)
        Globals.Ribbons.Ribbon1.lblProg.Label = str
        Globals.Ribbons.Ribbon1.lblProg.ShowLabel = show
    End Sub

    Public Sub doTranlate(ByVal fromLang As String, ByVal toLang As String, ByVal strCharset As String)
        bgHandler.CancelAsync() ' Stop current running
        setLabel(True, "Starting...")
        ' Prepare
        initTranslate()
        ' Clean up
        workingIndexes.Clear()
        ' Translate paramaters
        lngFrom = fromLang
        lngTo = toLang
        lngCharSet = strCharset
        ' Do async task
        bgHandler.RunWorkerAsync()
    End Sub

    Private Sub bgHandler_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        ' Keep looping to assign task among worker
        Do
            ' Update progress
            Dim percent As String = FormatPercent(workingIndexes.Count / curCells.addresses.Count, 1)
            setLabel(True, "Proceed (" & percent & ")")
            assignJob()
            ' If finished assign all job and no running tasks
            If workingIndexes.Count >= curCells.addresses.Count Then
                If Not (isBusy1 And isBusy2 And isBusy3 And isBusy4) Then
                    Application.Interactive = True
                    setLabel(False, "")
                    Exit Do
                End If
            End If
        Loop
    End Sub

    Private Sub assignJob()
        For i As Integer = 0 To curCells.addresses.Count - 1
            ' If not proceeded job
            If Not workingIndexes.Contains(i) Then
                Dim args As New cellArguments
                args.sheet = curCells.sheet
                args.address = curCells.addresses(i)
                args.value = curCells.values(i)
                If assignJobToWorker(args) Then
                    workingIndexes.Add(i)
                Else
                    Exit For
                End If
            End If
        Next
    End Sub

    ' Select non busy worker to assign task
    Private Function assignJobToWorker(ByVal args As cellArguments) As Boolean
        ' Assign job if not in busy state
        If Not isBusy1 Then
            doCellTranslate(args, 1)
            Return True
        End If

        If Not isBusy2 Then
            doCellTranslate(args, 2)
            Return True
        End If

        If Not isBusy3 Then
            doCellTranslate(args, 3)
            Return True
        End If

        If Not isBusy4 Then
            doCellTranslate(args, 4)
            Return True
        End If

        Return False
    End Function

    ' Translate and set cell value
    Private Sub doCellTranslate(ByVal cell As cellArguments, ByVal index As Integer)
        On Error GoTo NULL
        Select Case index
            Case 1
                isBusy1 = True
                translator1 = New cTranslator(cell.sheet, cell.address, cell.value)
                translator1.startTranslate(lngFrom, lngTo, lngCharSet)
            Case 2
                isBusy2 = True
                translator2 = New cTranslator(cell.sheet, cell.address, cell.value)
                translator2.startTranslate(lngFrom, lngTo, lngCharSet)
            Case 3
                isBusy3 = True
                translator3 = New cTranslator(cell.sheet, cell.address, cell.value)
                translator3.startTranslate(lngFrom, lngTo, lngCharSet)
            Case 4
                isBusy4 = True
                translator4 = New cTranslator(cell.sheet, cell.address, cell.value)
                translator4.startTranslate(lngFrom, lngTo, lngCharSet)
        End Select
NULL:
    End Sub

    ' Prepare all data before translate
    Private Sub initTranslate()
        Dim selRange As Range
        ' Check for excel selection
        Try
            Dim objRange As Range = Globals.ExcelTranslatorAddIn.Application.Selection
            ' Store original values to undo
            storeUndoValue(Application.ActiveSheet, objRange)
            selRange = objRange
        Catch ex As System.InvalidCastException
            ' Nothing to do
            Exit Sub
        End Try

        Application.Interactive = False
        ' Clean up 
        curCells = New cellProperties
        curCells.sheet = Application.ActiveSheet
        ' Get selection data
        For Each cell As Range In selRange
            If cell.Formula <> vbNullString Then ' if not empty cell
                If Not Left(cell.Formula, 1).Equals("=") Then ' if don't have formula
                    curCells.addresses.Add(cell.Address) ' Cell address
                    curCells.values.Add(cell.Value) ' Cell value
                End If
            End If
        Next
    End Sub

#Region "Undo transaction"

    ' Start undo transaction
    Public Sub doUndo()
        On Error Resume Next
        If undoCells.Count = 0 Then Exit Sub
        ' Protect sheet
        Application.Interactive = False
        ' Undo for lastest transaction
        With undoCells.Item(undoCells.Count - 1)
            For i As Integer = 0 To .addresses.Count - 1
                .sheet.range(.addresses(i)).Value = .values(i)
            Next
        End With
        ' Remove after undo
        undoCells.RemoveAt(undoCells.Count - 1)
        setUndoButtonState()

        Application.Interactive = True
    End Sub

    Private Sub storeUndoValue(ByVal sheet As Object, ByVal ranges As Range)
        On Error Resume Next
        ' No more than 10 undo
        Dim index As Integer = Math.Min(undoCells.Count - 1, 9)
        ' Remove last item if reaches 10 undo
        If index = 9 Then undoCells.RemoveAt(index)
        ' Create temp item
        Dim item As New cellProperties
        With item
            .sheet = sheet
            ' Get selection data
            For Each cell As Range In ranges
                .addresses.Add(cell.Address) ' Cell address
                .values.Add(cell.Value) ' Cell value
            Next
        End With
        ' Add to collection
        undoCells.Add(item)
        setUndoButtonState()
    End Sub

#End Region

    ' Initialize
    Private Sub ExcelTranslatorAddIn_Startup(sender As Object, e As EventArgs) Handles Me.Startup
        languageList = New List(Of languageProperties) ' Init collection
        ' Setup language list collection
        For Each str As String In languageListNames
            Dim tmp() As String = str.Split(";")
            ' Add to collection
            languageList.Add(New languageProperties(tmp(0), tmp(1), tmp(2)))
        Next

        ' Initial worker
        bgHandler.WorkerSupportsCancellation = True
        bgHandler.WorkerReportsProgress = True
        AddHandler bgHandler.DoWork, AddressOf bgHandler_DoWork

        ' Init ribbon
        Globals.Ribbons.Ribbon1.initRibbon()
    End Sub

    Private Sub setUndoButtonState()
        Globals.Ribbons.Ribbon1.btnUndo.Enabled = (undoCells.Count > 0)
    End Sub

    Private Sub translator1_Finished() Handles translator1.Finished
        ' Release worker busy task
        isBusy1 = False
    End Sub

    Private Sub translator2_Finished() Handles translator2.Finished
        ' Release worker busy task
        isBusy2 = False
    End Sub

    Private Sub translator3_Finished() Handles translator3.Finished
        ' Release worker busy task
        isBusy3 = False
    End Sub

    Private Sub translator4_Finished() Handles translator4.Finished
        ' Release worker busy task
        isBusy4 = False
    End Sub

End Class
