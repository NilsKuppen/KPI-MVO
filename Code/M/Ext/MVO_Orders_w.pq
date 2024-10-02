let

    Src = MVO_Afzet{0}, EFC = MVO_Afzet{1},

    // Function for sheet 3 + column with EFC#
    f = (i as number) as table => 
        let
            t = Table.PromoteHeaders( Src{i}[Data]{2} )
        in  Table.AddColumn( t, "EFC", each EFC{i} ),

    // Use function f with counter i to cycle through files and EFC#s
    g = List.Generate
    (
        () => [ a = f(0), i = 0 ],
        each [i]<List.Count(Src),
        each [ a = f(i), i = [i] +1 ],
        each [a]
    ),

    Output = Table.TransformColumns(Table.Combine( g ), {"DT", each Text.Remove( _, "-") } )
in
    Output