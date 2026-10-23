let
    Source = QueryHelper,
    ChangedType = Table.TransformColumnTypes(Source, {{"Qty", Int64.Type}})
in
    ChangedType
