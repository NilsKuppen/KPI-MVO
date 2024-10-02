let

    Src = MVO_Afzet{0}, EFC = MVO_Afzet{1},

    // Function to join Sheet1 and Sheet4 + column with EFC#
    f = (i as number) as table => 
        let
            t1 = Table.TransformColumnNames( Table.PromoteHeaders( Src{i}[Data]{0} ), (c) => Text.Remove(c, "-" ) ),
            t2 = Table.PromoteHeaders( Src{i}[Data]{3} ),
            t12 = Table.Join( t1, {"ARTNR"}, t2, {"ARTNR"} ),
            t = Table.AddColumn( t12, "EFC", each EFC{i} )
        in t,

    // Use function f with counter i to cycle through files and EFC#s
    g = List.Generate
    (
        () => [ a = f(0), i = 0 ],
        each [i]<List.Count(Src),
        each [ a = f(i), i = [i] +1 ],
        each [a]
    ),

    Output = Table.Combine( g )
in
    Output