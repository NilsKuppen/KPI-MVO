let
        // Maxvak data
    Src1 = Table.SelectColumns (
        SharePoint, { "ARTNR", "IUNITQYF", "EFC", "Dept", "LocTyp", "NAME", "Omd." }
    ),  // Sales
    Src2 = Table.SelectColumns (
        MVO_Afzet_w, {"ARTNR", "IUNITQYF", "EFC", "AANTAL" }
    ),
    Src = Table.Join ( 
        Src1, { "ARTNR", "IUNITQYF", "EFC" }, 
        Src2, { "ARTNR", "IUNITQYF", "EFC" } 
    ),
        // Add week + category 'Navultaken - '+location type
    Add.YW = Table.AddColumn( Src, "YW", each YW ),
    Add.Shelf = Table.AddColumn( 
        Add.YW, "CAT", 
        each "Navultaken - " & (
            if [#"Omd."] = "Y" then "Omdoos"    // Set 'omdoos' as seperate category
            else [LocTyp]                       
        )
     ),
        // Group and sum number of tasks
    Grp = Table.Group ( 
        Add.Shelf, { "YW", "EFC", "Dept", "CAT" }, { "VAL", each List.Sum([ AANTAL ]) }
    ),
        // Rename to conform to standard
    Tbl = Table.RenameColumns( Grp, { "Dept", "ZONE" } ),

    Output = Tbl
in

    Output