let
        // Maxvak data
    Src1 = Table.SelectColumns (
        SharePoint, { "ARTNR", "IUNITQYF", "EFC", "Dept", "LOCGRP.TYPE", "NAME" }
    ),  // Sales:
    Src2 = Table.SelectColumns (
        MVO_Afzet_w,
        let 
            l = Table.ColumnNames( MVO_Afzet_w ), 
            l1 = List.FirstN( l, 2 ),               // ARTNR, IUNITQYF
            l2 = List.Last( l )                     // EFC
        in 
            l1 & { l2 } & { YW }                    // Add week column by name YW
    ),
    Src = Table.Join ( 
        Src1, { "ARTNR", "IUNITQYF", "EFC" }, 
        Src2, { "ARTNR", "IUNITQYF", "EFC" } 
    ),

    Add.YW = Table.AddColumn( Src, "YW", each YW ),     // YYYYWW
    Add.Shelf = Table.AddColumn                         // Create category 'Afzet lvl' + shelf number #
    ( 
        Add.YW, "CAT",
        each "Afzet lvl - " & ( 
            if [LOCGRP.TYPE] <> "0" then "0"            // !Shelf == 0
            else Text.Range( [NAME], 4, 1 )             // else position 5 of location name
            ) 
    ),
        // Sum sales by week, EFC#, departement, category:
    Grp = Table.Group ( 
        Add.Shelf, { "YW", "EFC", "Dept", "CAT" }, { "VAL", each List.Sum( Record.Field( _, YW ) ) }
    ),
        // Rename to standard format:
    Tbl = Table.RenameColumns( Grp, { "Dept", "ZONE" } ),

    Output = Tbl
in

    Output