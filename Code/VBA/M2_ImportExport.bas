Attribute VB_Name = "M2_ImportExport"

'|*******************************************[ MAXVAK TOOL / Module: M2_ImportExport ]********************************************|
'|                                                                                                                                |
'|                                               [[[    Author: Nils Kuppen    ]]]                                                |
'|                                               [[[       For: JUMBO.com      ]]]                                                |
'|                                               [[[ EFC Den Bosch \ SiSu \ VO ]]]                                                |
'|                                                                                                                                |
'|                  This module contains all macro's that import or export source data necessary for data queries                 |
'|                                                                                                                                |
'|                      Relative file path: \\Code\VBA\M2_ImportExport                                                            |
'|                      Updated by fn: ThisWorkbook.update_vba                                                                    |
'|                      Triggered by: config.xml\config\update\vba = true                                                         |
'|                                                                                                                                |
'|**********************************************************[ (c)  2024 ]*********************************************************|



    '// export_KPI_sp( )
    '// Update weekly KPI's and export to SharePoint
    '// Source file is opened and table contents copied to first free line

Sub export_KPI_sp()

            Blad9.ListObjects(1).Range.Copy

    Set wb = Workbooks.Add

        With wb
                .Sheets(1).Range("A1").PasteSpecial xlPasteValues
                .SaveAs SharePoint & "KPI_MVO_Bron.xlsx"
                .Close
        End With

End Sub
