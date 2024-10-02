let
    // Central Online MVO Input files // EFC hours
    Src = Excel.Workbook(File.Contents(Central & "\2_Uren\Uren Overzicht EFCs.xlsx"), null, true),
    Tbl = Src{0}[Data], // Sheet 1
    H = Table.PromoteHeaders(Tbl, [PromoteAllScalars=true]),
    // Add YYYYWW format
    YW = Table.TransformColumns (
        H, {
            {
                "Datum", each 
                    let // x = date, y = year , weeknr in "00" format
                        x = Date.From(_), 
                        y = Text.From( Date.Year(x) ), 
                        w = Text.End( "0" & Text.From( Date.WeekOfYear(x) ), 2 )
                    in y & w
            },
            {
                "EFC", each Text.From(_)
            }
        }
    )
in
    YW