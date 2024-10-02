let
    Src = SharePoint,
    Grp = Table.Group   // Group by EFC, departement + category, week + count C state
    (
        Src, { "EFC", "Dept" },
        {
            { "CAT", each "Loc - Closing" },
            { "YW", each YW },
            { "VAL", each List.Count( List.Select( [STATE], each _="C" ) ), type number }
        }
    ),  // Conform to standard:
    Buf = /*Table.Buffer(*/ Grp /*)*/ ,
    Tbl = Table.FromColumns
    (
        { Buf[YW], Buf[EFC], Buf[Dept], Buf[CAT], Buf[VAL] },
        { "YW", "EFC", "ZONE", "CAT", "VAL" }
    ),
    
    Output = Tbl
in
    Output