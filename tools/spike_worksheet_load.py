"""Disposable Excel probe for loading an inline static Power Query into a worksheet."""
from __future__ import annotations

import tempfile
from pathlib import Path

import pythoncom
import win32com.client


def main():
    pythoncom.CoInitialize()
    excel = workbook = reopened = None
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-sheet-spike-') as temp:
            path = Path(temp) / 'worksheet.xlsx'
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            print(f'Excel {excel.Version} build {excel.Build}')
            workbook = excel.Workbooks.Add()
            name = 'PQM_Sheet_Spike'
            formula = 'let Source = #table({"Value"}, {{1}, {2}}) in Source'
            workbook.Queries.Add(name, formula)
            connection_string = (f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                                 f'Location={name};Extended Properties=""')
            sheet = workbook.Worksheets.Item(1)
            list_object = sheet.ListObjects.Add(
                SourceType=3,
                Source=connection_string,
                LinkSource=True,
                XlListObjectHasHeaders=1,
                Destination=sheet.Range('B3'),
            )
            query_table = list_object.QueryTable
            query_table.CommandType = 2
            query_table.CommandText = f'SELECT * FROM [{name}]'
            query_table.BackgroundQuery = False
            result = query_table.Refresh(False)
            query_table.WorkbookConnection.Name = f'Query - {name}'
            print('refresh result:', result)
            print('query tables:', sheet.QueryTables.Count)
            print('list objects:', sheet.ListObjects.Count)
            print('result rows:', query_table.ResultRange.Rows.Count)
            print('result address:', query_table.ResultRange.Address)
            print('cell values:', sheet.Range('B3:B5').Value)
            workbook.SaveAs(str(path), 51)
            workbook.Close(False)
            workbook = None
            reopened = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True,
                                            IgnoreReadOnlyRecommended=True, AddToMru=False)
            sheet = reopened.Worksheets.Item(1)
            table = sheet.ListObjects.Item(1)
            assert reopened.Queries.Item(name).Formula == formula
            print('reopened table connection:', table.QueryTable.WorkbookConnection.Name)
            print('workbook connections:', [reopened.Connections.Item(i).Name for i in range(1, reopened.Connections.Count + 1)])
            print('table connection string:', table.QueryTable.WorkbookConnection.OLEDBConnection.Connection)
            assert table.Range.Address == '$B$3:$B$5'
            assert tuple(value[0] for value in sheet.Range('B3:B5').Value) == ('Value', 1.0, 2.0)
            print('reopen: formula, connection, table range, and cell values verified')
            reopened.Close(False)
            reopened = None
        return 0
    finally:
        for item in (reopened, workbook):
            if item is not None:
                try:
                    item.Close(False)
                except Exception:
                    pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
