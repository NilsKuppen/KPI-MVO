let
        // Select Maxvak data with active picklocations
    Src = Table.SelectColumns( 
        Table.SelectRows( 
            SharePoint, 
            each [ArtNr] <> "" 
            and [ArtNr] <> null
            and [STATE] = "A" 
        ), 
        { "EFC", "Dept", "d Grp" } 
    ),  // Group by EFC, department + count rows with correct group over total rows:
    Grp = Table.Group
    (
        Src, { "EFC", "Dept" },
        { "VAL", each Table.RowCount( Table.SelectRows( _, each [ d Grp ] = null ) ) / Table.RowCount(_) }
    ),  // Add category 'Slotting - Grp%' and conform to standard:
    Tbl = Table.FromRecords
    (
        Table.TransformRows (
            Grp, 
            (r) => [ YW=YW, EFC=r[EFC], ZONE=r[Dept], CAT = "Slotting - Grp %", VAL=r[VAL] ]
        )
    ),

    Output = Tbl

in

    Output