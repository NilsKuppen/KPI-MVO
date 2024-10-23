let

    Src = MTO_UrenOverzicht,            // Hours
    YW.last = List.Max( Src[Datum] ),   // Last week
        // Filter table on last week and refilling
    Fltr = Table.SelectRows( Src, each [Datum] = YW.last and Text.Contains( [#"Afdeling (versimpeld)"], "Navul" ) ),
    Ins = Table.FromRecords
    (
        {
            [ EFC="8833", Datum=YW.last, Zone="Ongekoeld",  #"Som van Geklokt"=0 ],
            [ EFC="8833", Datum=YW.last, Zone="Vers",       #"Som van Geklokt"=0 ],
            [ EFC="8833", Datum=YW.last, Zone="Diepvries",  #"Som van Geklokt"=0 ],
            [ EFC="8836", Datum=YW.last, Zone="Ongekoeld",  #"Som van Geklokt"=0 ],
            [ EFC="8836", Datum=YW.last, Zone="Vers",       #"Som van Geklokt"=0 ],
            [ EFC="8836", Datum=YW.last, Zone="Diepvries",  #"Som van Geklokt"=0 ],
            [ EFC="8837", Datum=YW.last, Zone="Ongekoeld",  #"Som van Geklokt"=0 ],
            [ EFC="8837", Datum=YW.last, Zone="Vers",       #"Som van Geklokt"=0 ],
            [ EFC="8837", Datum=YW.last, Zone="Diepvries",  #"Som van Geklokt"=0 ]
        }    
    ),
    Tbl = Table.Combine({ Fltr, Ins }),
        // Group + sum hours
    Grp = Table.Group( Tbl, { "EFC", "Datum", "Zone" }, {{ "Uren", each List.Sum([Som van Geklokt]), type number }} ),
        // Use LuTs to set departement
    Join.Lut = Table.NestedJoin( Grp, { "Zone" }, Lut_Zone, { "Cat" }, "Lut" ), Xp.Lut = Table.ExpandTableColumn( Join.Lut, "Lut", { "Descr"} ), 

    t = Table.Distinct( Xp.Lut ), 
        // Conform to standard
    Sel = Table.FromColumns
    ( 
        { t[Datum], t[EFC], t[Descr], List.Repeat( { "Uren - Navullen" }, Table.RowCount(t) ), t[Uren] }, 
        { "YW", "EFC", "ZONE", "CAT", "VAL" } 
    ),

    Output = Sel
in
    Output