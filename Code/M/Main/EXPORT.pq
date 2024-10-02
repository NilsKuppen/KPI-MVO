    // Filter out new week in last KPI export and combine with with new data 
Table.Combine({
    Table.SelectRows( KPI_Old, each [YW] <> YW ),
    KPI_New
})