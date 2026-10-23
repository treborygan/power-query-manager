"""Disposable Excel probe for a Power Query connection loaded into the Data Model."""
from __future__ import annotations

import tempfile
from pathlib import Path

import pythoncom
import win32com.client


def main():
    pythoncom.CoInitialize()
    excel = workbook = reopened = None
    try:
        with tempfile.TemporaryDirectory(prefix='pqm-model-spike-') as temp:
            path = Path(temp) / 'model.xlsx'
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            print(f'Excel {excel.Version} build {excel.Build}')
            workbook = excel.Workbooks.Add()
            name = 'PQM_Model_Spike'
            formula = 'let Source = #table({"Value"}, {{1}, {2}}) in Source'
            workbook.Queries.Add(name, formula)
            connection_string = (f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                                 f'Location={name};Extended Properties=""')
            connection = workbook.Connections.Add2(
                f'Query - {name}', 'Disposable Data Model spike', connection_string,
                f'SELECT * FROM [{name}]', 2, True, False,
            )
            print('InModel before refresh:', connection.InModel)
            connection.Refresh()
            print('InModel after refresh:', connection.InModel)
            print('Model table count:', workbook.Model.ModelTables.Count)
            for index in range(1, int(workbook.Model.ModelTables.Count) + 1):
                model_table = workbook.Model.ModelTables.Item(index)
                print('model table:', model_table.Name, 'records:', model_table.RecordCount,
                      'source:', model_table.SourceWorkbookConnection.Name)
            workbook.SaveAs(str(path), 51)
            workbook.Close(False)
            workbook = None
            reopened = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True,
                                            IgnoreReadOnlyRecommended=True, AddToMru=False)
            print('Reopened InModel:', reopened.Connections.Item(f'Query - {name}').InModel)
            print('Reopened model table count:', reopened.Model.ModelTables.Count)
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
