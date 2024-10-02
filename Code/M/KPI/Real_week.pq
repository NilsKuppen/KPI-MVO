let 
        // Get order/CE amounts per week and filter by target week
    Src = Table.SelectRows( MVO_Orders_w, each [DT] = YW ),
        // Transform records to create missing columns and conform to standard:
    t = Table.FromRecords 
    (
        Table.TransformRows
        (
            Src, (r) =>
            [
                YW = r[DT], EFC = r[EFC], ZONE = null,
                T = #table ( 
                    {"CAT", "VAL"}, 
                    { { "Week orders", r[ORDERS] }, { "Week CE", r[CE] } } 
                )
            ]
        )
    ),
    x = Table.ExpandTableColumn( t, "T", { "CAT", "VAL" } ),
    
    Output = x

in

    Output