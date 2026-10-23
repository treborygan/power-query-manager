let
    Source = Table.FromRecords({
        [Part_ID = "KRC100000/1 R1A", Location = "FLTY", Qty = 3],
        [Part_ID = "KRC100001/1 R1A", Location = "BUFR", Qty = 5]
    })
in
    Source
