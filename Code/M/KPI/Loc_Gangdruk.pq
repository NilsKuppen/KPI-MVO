let
    Src1 = SharePoint,
    Src2 = Table.SelectColumns (
        MVO_Afzet_w,
        let 
            l = Table.ColumnNames( MVO_Afzet_w ), 
            l1 = List.FirstN( l, 2 ),               // ARTNR, IUNITQYF
            l2 = List.Last( l )                     // EFC
        in 
            l1 & { l2 } & { YW }                    // + column by week
    ),
    Src = Table.Join ( 
        Src1, { "ARTNR", "IUNITQYF", "EFC" }, 
        Src2, { "ARTNR", "IUNITQYF", "EFC" } 
    ),
    Grp = Table.Group (     // Group by EFC, departement, corridor + week + sum sales over corridor length
        Src, { "EFC", "Dept", "GANG" },
        {
            { "YW", each YW },
            { 
                "VAL", 
                each Number.Round (
                    List.Sum( Record.Field( _, YW ) ) / List.First([ GANG_LENGTE ])
                ), Int64.Type 
            }
        }
    ),  // Group records
    Grp.Tbl = Table.Group ( Grp, { "YW", "EFC", "Dept" }, {{ "GROUP", each _, type table }} ),
        // Combine corridor + value to VAL + add category 'Gangdruk -'+rank
    Rank = Table.TransformColumns (
        Grp.Tbl,
        {
            "GROUP", 
            each
            Table.FromRecords ( 
                Table.TransformRows (
                    Table.FirstN (
                        Table.AddIndexColumn ( 
                            Table.Sort(_, { "VAL", Order.Descending} ), 
                            "CAT", 1
                        ), 5
                    ), 
                    (r) => 
                    Record.TransformFields ( 
                        r, {
                            { "CAT", each "Gangdruk - " & Text.From(_) },   // create category, use index as ranking
                            { "VAL", each r[GANG] & ", " & Text.From(_) }   // delim corrider + value with ', '
                        }
                    )
                )
            )
        }
    ),
        // Expand new records to columns and set to standard format:
    xp = Table.ExpandTableColumn ( Rank, "GROUP", { "CAT", "VAL" } ),
    Tbl = Table.FromColumns (
        { xp[YW], xp[EFC], xp[Dept], xp[CAT], xp[VAL] }, 
        { "YW", "EFC", "ZONE", "CAT", "VAL" }
    ),

    Output = Tbl
in
    Output