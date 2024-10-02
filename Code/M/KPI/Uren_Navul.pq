let

    Src = MTO_UrenOverzicht,            // Hours
    YW.last = List.Max( Src[Datum] ),   // Last week
        // Filter table on last week and refilling
    Fltr = Table.SelectRows( Src, each [Datum] = YW.last and Text.Contains( [#"Afdeling (versimpeld)"], "Navul" ) ),
        // Group + sum hours
    Grp = Table.Group( Fltr, { "EFC", "Datum", "Zone", "Afdeling (versimpeld)" }, {{ "Uren", each List.Sum([Som van Geklokt]), type number }} ),
        // Use LuTs to set departement
    Join.Lut = Table.NestedJoin( Grp, { "Zone" }, Lut_Zone, { "Cat" }, "Lut" ), Exp.Lut = Table.ExpandTableColumn( Join.Lut, "Lut", { "Descr" } ),

    t = Table.Distinct( Exp.Lut ), 
        // Conform to standard
    Sel = Table.FromColumns
    ( 
        { t[Datum], t[EFC], t[Descr], List.Transform( t[Descr], each "Uren - Navullen" ), t[Uren] }, 
        { "YW", "EFC", "ZONE", "CAT", "VAL" } 
    ),

    Output = Sel
in
    Output