let

    Src = Table.SelectColumns( SharePoint, { "EFC", "Dept", "LocTyp", "STATE" } ),
        // Group by EFC, departement, location type + count empty locations over total locations
    Grp = Table.Group
    (
        Src, { "EFC", "Dept", "LocTyp" },
        { "VAL", each Table.RowCount( Table.SelectRows( _, each [ STATE ] <> null ) ) / Table.RowCount(_) }
    ),  // Add category 'Loc - vol % -'+location type and conform to standard:
    Tbl = Table.FromRecords
    (
        Table.TransformRows (
            Grp, 
            (r) => [ YW=YW, EFC=r[EFC], ZONE=r[Dept], CAT = "Loc - Vol % - " & r[LocTyp], VAL=r[VAL] ]
        )
    ),

    Output = Tbl

in

    Output