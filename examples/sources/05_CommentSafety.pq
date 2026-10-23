let
    Source = Table.FromRecords({[A = 1]}),
    // MissingQueryInComment should NOT become a dependency.
    Marker = "QueryHelper is just text; not a dependency",
    Result = Table.AddColumn(Source, "Marker", each Marker)
in
    Result
