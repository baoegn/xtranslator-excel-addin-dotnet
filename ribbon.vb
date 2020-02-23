Imports System.ComponentModel
Imports Microsoft.Office.Tools.Ribbon

Public Class ribbon
    ' Initialize
    Public Sub initRibbon()
        ' Clean up
        ddFromLang.Items.Clear()
        ddToLang.Items.Clear()
        ' Load item to dropdown
        For Each language As languageProperties In languageList
            ' Create new dropdown items for 2 dropdowns
            Dim item() As RibbonDropDownItem = {Factory.CreateRibbonDropDownItem,
                                                Factory.CreateRibbonDropDownItem}
            ' Set label
            item(0).Label = language.name : item(1).Label = language.name
            ' Add to dropdown
            ddFromLang.Items.Add(item(0))
            ddToLang.Items.Add(item(1))
        Next
        ' Default from settings
        ddFromLang.SelectedItemIndex = My.Settings.fromLanguage
        ddToLang.SelectedItemIndex = My.Settings.toLanguage
    End Sub

    Private Sub btnSwitch_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSwitch.Click
        ' Switch between language
        Dim tmpIndex As Integer
        tmpIndex = ddToLang.SelectedItemIndex
        ddToLang.SelectedItemIndex = ddFromLang.SelectedItemIndex
        ddFromLang.SelectedItemIndex = tmpIndex
    End Sub

    Private Sub btnTranslate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnTranslate.Click

        ' In case no internet connection then exit
        If Not hasInternet() Then
            MsgBox("No Internet connection! Please check!", vbExclamation + vbOKOnly, "No Internet")
            Exit Sub
        End If

        ' Get current dropdowns index
        Dim fromIndex = ddFromLang.SelectedItemIndex
        Dim toIndex = ddToLang.SelectedItemIndex

        ' Get language code & charset of selected item
        Dim lngFrom As String = languageList(fromIndex).code
        Dim lngTo As String = languageList(toIndex).code
        Dim lngCharset As String = languageList(toIndex).charset

        ' Save to settings
        My.Settings.fromLanguage = ddFromLang.SelectedItemIndex
        My.Settings.toLanguage = ddToLang.SelectedItemIndex
        My.Settings.Save()

        ' Do translate
        Globals.ExcelTranslatorAddIn.doTranlate(lngFrom, lngTo, lngCharset)
    End Sub

    Private Sub btnAbout_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAbout.Click
        MsgBox("This is non-commercial add-in for Excel using HTTP request to Google Translate." & vbNewLine & vbNewLine &
                "Icon credits to:" & vbNewLine & " - https://www.flaticon.com" & vbNewLine &
                " - https://ionicons.com/" & vbNewLine & vbNewLine & "Thank for using!" & vbNewLine & "------------------------------------------------------------" & vbNewLine &
                "Written by: baoegn@gmail.com to Janis" & vbNewLine, vbInformation + vbOKOnly, "About xTranslator")
    End Sub

    Private Sub btnUndo_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUndo.Click
        Globals.ExcelTranslatorAddIn.doUndo()
    End Sub

End Class
