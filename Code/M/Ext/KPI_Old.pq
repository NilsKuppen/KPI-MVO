let // Connect to sharepoint:
    Src = SharePoint.Files("https://jumbosupermarkten.sharepoint.com/sites/CentralOnlineMVO/", [ApiVersion = 15]),
    // Open last KPI export:
    File = Src
        {
            [
                Name="KPI_VO_Bron.xlsx",
                #"Folder Path"="https://jumbosupermarkten.sharepoint.com/sites/CentralOnlineMVO/Gedeelde documenten/Slotting/3. Exports/"
            ]
        }[Content],
    wb = Excel.Workbook(File)[Data]{0},
    H = Table.PromoteHeaders( wb ),

    Output = Table.TransformColumnTypes(H,{{"YW", type text}, {"EFC", type text}})
in
    Output