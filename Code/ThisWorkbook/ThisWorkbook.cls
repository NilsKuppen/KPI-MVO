VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const cfg = "Code\config.xml"
Const pad = "O:\Map A-F\Central Online MVO\1_Tool & Dashboard\KPI MVO\"

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ThisWorkbook.Saved = True

End Sub

Private Sub force_m_update()

update_pq (True)

End Sub

Private Sub Workbook_Open()

If Application.Visible = False Then Exit Sub

    Dim rTxt As Range: Set rTxt = Range("LOAD_TXT")
    Dim m As Boolean, vba As Boolean
    
        ThisWorkbook.Queries.FastCombine = True
        Application.DisplayAlerts = False
        
        Range("PATH") = ThisWorkbook.Path
        
            'Blad3.Visible = xlSheetVisible
            Blad3.Select
            
                i% = 0: c% = 6: x% = 0
                load_bar i, c: rTxt = "Opstarten..."
            
        DoEvents
        
            If ThisWorkbook.ReadOnly Then
                load_bar i, c: rTxt = "Read-Only: kan niet updaten"
                GoTo Eind
            End If
            
            load_bar i, c: rTxt = "Updates controleren..."
        
            If Range("version") <> version_check() Then
                    c = c + 1
                    load_bar i, c: rTxt = "Nieuwe versie beschikbaar"
                    update_check m, vba
            End If
            
        DoEvents
            load_bar i, c: rTxt = "M-Code updaten..."
        
            update_pq (m)

        DoEvents
            load_bar i, c: rTxt = "Bestand up-to-date"
            
            If Range("version") <> version_check() Then
                    append_log
                    Range("version") = version_check()
            End If
Eind:
            DoEvents
                load_bar i, c: rTxt = "Verbinding maken..."

            DoEvents
                load_bar 1, 1: rTxt = "Klaar!"

            Blad6.Select
                load_bar 0, 10: rTxt = ""
            'Blad3.Visible = xlSheetVeryHidden
            
        If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
                
        Application.DisplayAlerts = True
        
Exit Sub

errHndlr:

    error_log
    MsgBox "Fout: opstarten onderbroken. Neem contact op met Nils Kuppen." & vbCrLf & vbCrLf & "Error:" & Err.Description

End Sub

Public Sub load_bar(i%, c%)

    Dim r As Range
    
        Set r = Range("LOAD_BAR")
            max% = r.Cells.Count
            x% = -Int(-(i / c * max))
        If i = 0 Then r.Cells.Interior.Color = xlNone
        If i > 0 Then
            r.Resize(1, x).Interior.Color = RGB(0, 255, 0)
        End If
            i = i + 1
           
End Sub

Private Sub force_vba()

    Dim FSO As Object
    Dim folder As Object
    Dim file As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")

        Set folder = FSO.getfolder(pad & "Code\VBA\")
    
            With ThisWorkbook.VBProject
                '//Start delete loop
                    c% = .VBComponents.Count: i% = 0
                    
                    load_bar i, c
                On Error Resume Next
                    For Each VBComp In .VBComponents
                            .VBComponents.Remove VBComp
                            DoEvents
                        
                        load_bar i, c

                    Next VBComp
                On Error GoTo 0
                
                '//End delete loop

                c = folder.files.Count: i = 0
                load_bar i, c

                '//Start add loop
                    For Each file In folder.files
                            .VBComponents.Import (file.Path)
                            load_bar i, c
                    Next file
                '//End add loop
            End With '// VBProject
            
        i = 0: load_bar i, c

End Sub

Private Sub update_pq(x As Boolean) '//x=true forces update

    Dim FSO     As Object
    Dim folder  As Object
    Dim folders As Object
    Dim files   As Object
    Dim file    As Object
    Dim m       As String '//M-Code
    Dim i%, c%

    Set FSO = CreateObject("Scripting.FileSystemObject")

    Set folder = FSO.getfolder(ThisWorkbook.Path & "\Code\m\")
    Set folders = folder.subfolders

            c = folder.files.Count
        For Each folder In folders
            c = c + folder.files.Count
        Next folder
            i = 0: load_bar i, c
            '//Loop through folders\files
            For Each folder In folders
                For Each file In folder.files
                
                    dt = DateDiff("s", file.datelastmodified, FileDateTime(ThisWorkbook.FullName))
                
                    If dt < 0 Or x = True Then
                        n = Left(file.Name, InStr(file.Name, ".pq") - 1)
                        m = ""
                        
                        hnd = FreeFile
                                    
                        Open file.Path For Input As hnd
                            Do Until EOF(hnd)
                                Line Input #hnd, Row
                                m = m & Row & vbNewLine
                            Loop
                        Close #hnd
                        
                        On Error Resume Next
                            ThisWorkbook.Queries.Add n, m
                        On Error GoTo 0
                            ThisWorkbook.Queries.Item(n).Formula = m
                    End If

                    load_bar i, c

                Next file
            
            Next folder
            
            i = 0: load_bar i, c
            
            If x = True Then del_pq
            
End Sub

Private Function version_check() As String

    Dim xmlDoc As New DOMDocument30

        xmlDoc.Load (pad & cfg)
        xmlDoc.async = False
            
    Dim mainNode As IXMLDOMNode

    Set mainNode = xmlDoc.SelectSingleNode("//version")

        version_check = mainNode.Text
        
End Function

Private Sub update_check(ByRef m As Boolean, ByRef vba As Boolean)

    Dim xmlDoc As New DOMDocument30
    
        xmlDoc.Load (pad & cfg)
        xmlDoc.async = False
            
    Dim mainNode As IXMLDOMNode

    Set mainNode = xmlDoc.SelectSingleNode("//update/m")

        m = (mainNode.Text = "true")

    Set mainNode = xmlDoc.SelectSingleNode("//update/vba")

        vba = (mainNode.Text = "true")
        
End Sub

Private Sub append_log()

    Dim xmlDoc As New DOMDocument30

        xmlDoc.Load (pad & cfg)
        xmlDoc.async = False
            
    Dim mainNode As IXMLDOMNode

    Set c = xmlDoc.SelectSingleNode("//update/comment")
    Set u = xmlDoc.SelectSingleNode("//update/user")
    Set d = xmlDoc.SelectSingleNode("//update/dt")
    Set v = xmlDoc.SelectSingleNode("//version")

        Open ThisWorkbook.Path & "\changelog.txt" For Append As #1
        Print #1, "# [" & d.Text & "] "; v.Text & " : "; c.Text; " [" & u.Text & "]"
        Close #1

End Sub

Private Sub error_log()

    Dim xmlDoc As New DOMDocument30

        xmlDoc.Load (pad & cfg)
        xmlDoc.async = False
            
    Dim mainNode As IXMLDOMNode

        c = "Error: " & Err.Number & " - " & Err.Description
        u = Application.UserName
        d = Date
    Set v = xmlDoc.SelectSingleNode("//version")

        Open ThisWorkbook.Path & "\changelog.txt" For Append As #1
        Print #1, "# [" & d & "] "; v.Text & " : "; c; " [" & u & "]"
        Close #1

End Sub

Private Sub del_pq()


Dim FSO As Object
Dim L As Object
Dim m As Object

Set FSO = CreateObject("Scripting.FileSystemObject")
Set L = CreateObject("Scripting.Dictionary")

Set m = FSO.getfolder(ThisWorkbook.Path & "\Code\M\")

    For Each folder In m.subfolders

        For Each file In folder.files

            L.Add FSO.GetBaseName(file), 1
        
        Next file

    Next folder
    
    c% = ThisWorkbook.Queries.Count

    For Each Query In ThisWorkbook.Queries
    
        load_bar i%, c
    
        If Not L.Exists(Query.Name) Then ThisWorkbook.Queries(Query.Name).Delete
        
    Next Query

End Sub


