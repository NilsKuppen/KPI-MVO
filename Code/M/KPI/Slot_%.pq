let
        // Get maxvak data and select active shelf locations 
    Src = Table.SelectColumns ( SharePoint, { "ARTNR", "IUNITQYF", "EFC", "Dept", "LOCGRP.TYPE", "% {New, Old}", "STATE" } ),
    Schap = Table.SelectRows( Src, each [LOCGRP.TYPE] = "0" and [STATE] = "A" ),
        // Add week and category columns
    YW = Table.AddColumn( Schap, "YW", each YW ),
    Cat = Table.AddColumn
    ( 
        YW, "CAT",
        each 
            let x = [#"% {New, Old}"]                   // Slotting percentage
            in "Slotting % - " & (                      // Category +
                if x < 0.8 or x = null then "< 80%"     // Sub: Low
                else if x > 1.2 then "> 120%"           // Sub: High
                else "> 80 < 120%"                       // Sub: Within Margin
            )
    ),
        // Group and count rows:
    Grp = Table.Group
    (
        Cat, { "YW", "EFC", "Dept", "CAT" },
        { "VAL", Table.RowCount }
    ),
        // Rename to standard:
    Tbl = Table.RenameColumns( Grp, { "Dept", "ZONE" } ),

    Output = Tbl

in

    Output