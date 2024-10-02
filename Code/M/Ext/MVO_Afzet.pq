let

        // Look at all files in folder
    Src = Folder.Files( Central ),
        // Filter on pattern
    Files = Table.Sort( Table.SelectRows( Src, each Text.EndsWith( [Name], "_Afzet X-12wkn.xlsx" ) ), { "Name", Order.Ascending } ),
        // Open all Workbooks
    File = List.Transform( Files[Content], each Excel.Workbook(_) ),
        // List all EFC's
    EFC = List.Transform( Files[Name], each Text.Start( _, 4 ) )

in

    { File, EFC }