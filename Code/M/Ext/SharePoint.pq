let // Connect to sharepoint:
    Src = SharePoint.Files("https://jumbosupermarkten.sharepoint.com/Sites/CentralOnlineMVO/", [ApiVersion = 15]),
    Files = Table.SelectRows(Src, each Text.StartsWith([Name], "Maxvak_Tabel_")),   // Find all Maxvak exports
    // Function to prepare tables to combine/join:
    fn.File = (t as table, i as number, n as text) as table => 
        let A = Table.PromoteHeaders( t{i}[Data] ), 
            B = Table.AddColumn( A, "EFC", each n, type text )
        in B,
    // 8833:
    File_0 = Excel.Workbook(Files{0}[Content]),
    File_0_0 = fn.File( File_0, 0, "8833" ),
    File_0_1 = fn.File( File_0, 1, "8833" ),
    // 8836:
    File_1 = Excel.Workbook(Files{1}[Content]),
    File_1_0 = fn.File( File_1, 0, "8836" ),
    File_1_1 = fn.File( File_1, 1, "8836" ),
    // 8837:
    File_2 = Excel.Workbook(Files{2}[Content]),
    File_2_0 = fn.File( File_2, 0, "8837" ),
    File_2_1 = fn.File( File_2, 1, "8837" ),
    // Combine all same sheets:
    Tbl =
    { 
        Table.Combine( { File_0_0, File_1_0, File_2_0 } ),  // All location overview sheets
        Table.Combine( { File_0_1, File_1_1, File_2_1 } )   // All Maxvak overview sheets
    },
    // Join location sheets with maxvak sheets for one table
    Join.Tbl = Table.Join( 
        Tbl{0}, { "EFC", "NAME" }, 
        Table.PrefixColumns( Tbl{1}, "mv" ),
         { "mv.EFC", "mv.Loc" }, JoinKind.LeftOuter 
    ),
    f2.Col = Table.TransformColumnNames( 
        Join.Tbl, (c) => if Text.StartsWith( c, "mv.") then Text.Range( c, 3 ) else c 
    ),
    // Convert zones to departments and locationtypes from number to description using lookup tables
    // Zones:
    Lut.Key = Table.AddColumn( f2.Col, "Key", each Text.Start( [LOCGRP.G ], 1 ) ),
    Join.Lut = Table.NestedJoin( Lut.Key, "Key", Lut_Zone, "Type", "LUT" ),
    Exp.Lut = Table.ExpandTableColumn( Join.Lut, "LUT", {"Descr"}, {"Dept"} ),
    // Types:
    Join.Lut.2 = Table.NestedJoin( Exp.Lut, "LOCGRP.TYPE", Lut_Type, "Type", "LUT" ),
    Exp.Lut.2 = Table.ExpandTableColumn( Join.Lut.2, "LUT", {"Descr"}, {"LocTyp"} ),   

    Output = Exp.Lut.2
in 
    Output