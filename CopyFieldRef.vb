Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE90
Imports EnvDTE90a
Imports EnvDTE100
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Text

Public Module CopyFieldRefVer2

    Private clipText As String

    Sub CopyFieldRefVer2()
        Dim textSelection As TextSelection = DTE.ActiveDocument.Selection
        Dim reg As Regex = New Regex("\b(ID|Name)=""([^""]*)""", RegexOptions.None)
        Dim m As Match = reg.Match(textSelection.Text)
        Dim previouseOne As Match = Nothing

        Dim id As String = String.Empty
        Dim name As String = String.Empty
        Dim idCount As Integer = 0
        Dim nameCount As Integer = 0

        Dim oneLiner As Boolean = False 'set this to "True" if you want refs to be on one line
        Dim tab As String = vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab
        Dim space As String = " "
        Dim separator As String = If(oneLiner, space, tab)

        Dim result As New StringBuilder

        While m.Success
            If m.Groups(1).Value().Equals("ID") AndAlso String.IsNullOrEmpty(id) Then
                id = vbCrLf + "<FieldRef " + m.Groups(0).Value()
                idCount += 1
            ElseIf m.Groups(1).Value().Equals("Name") AndAlso String.IsNullOrEmpty(name) Then
                name = separator + m.Groups(0).Value() + " />"
                nameCount += 1
            End If

            If Not String.IsNullOrEmpty(id) AndAlso Not String.IsNullOrEmpty(name) Then
                result.Append(id + name)
                id = String.Empty
                name = String.Empty
            End If

            previouseOne = m
            m = m.NextMatch()
        End While

        If Not idCount = nameCount Then
            result.Insert(0, "<!-- ERROR 9: uneven number of IDs and names. -->" + vbNewLine + "<!-- Please check your selection/code you monkey! -->")
        End If

        clipText = result.ToString()
        RunThread(AddressOf CopyToClipboard)
    End Sub

    Private Function RunThread(ByVal fct As Threading.ThreadStart)
        Dim thread As New Threading.Thread(fct)
        thread.ApartmentState = Threading.ApartmentState.STA

        thread.Start()
        thread.Join()
    End Function

    Private Sub CopyToClipboard()
        My.Computer.Clipboard.SetText(clipText)
    End Sub
End Module
