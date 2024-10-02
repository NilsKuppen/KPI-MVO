let
    query.list = {
        Uren_Navul,
        Loc_Closing,
        Loc_Gangdruk,
        Slot_Grp,
        Loc_Vol,
        Afzet_Lvl,
        Taken_Navul,
        #"Slot_%",
        Real_week
    }
in
    Table.Combine( query.list )