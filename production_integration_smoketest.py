from pathlib import Path
import tempfile
import tkinter as tk
from tkinter import ttk
from app import Database, OfflineDatabaseApp

source = Path(__file__).resolve().parent / 'initial_data.xlsx'
with tempfile.TemporaryDirectory(prefix='production_integration_') as temp_dir:
    db = Database(Path(temp_dir))
    db.import_excel(source, replace=True)
    db.create_production_record(1, '2026-08-21', '1,200', '1,180', '林志強製作廠', '包裝線 A，檢驗合格後入庫。')
    db.create_production_record(1, '2026-07-05', '900', '900', '陳建宏製作廠', '')

    app = OfflineDatabaseApp(db)
    app.update_idletasks()
    fields = db.list_fields()
    brand = next(f for f in fields if f['system_key'] == 'brand_name')
    production_cols = [c for c in app.tree['columns'] if str(c) == f"production{int(brand['id'])}"]
    assert production_cols == [f"production{int(brand['id'])}"], production_cols
    assert app.tree.heading(production_cols[0], 'text') == ''
    assert int(app.tree.column(production_cols[0], 'width')) == 34

    app.open_production_records(1)
    app.update_idletasks()
    wins = [child for child in app.winfo_children() if isinstance(child, tk.Toplevel) and child.title().startswith('生產記錄 - ')]
    assert len(wins) == 1, len(wins)
    win = wins[0]
    trees = []
    def walk(widget):
        for child in widget.winfo_children():
            if isinstance(child, ttk.Treeview):
                trees.append(child)
            walk(child)
    walk(win)
    assert trees, 'production history table missing'
    assert tuple(trees[0]['columns']) == ('seq', 'date', 'order', 'stock', 'manufacturer', 'remark', 'actions')
    assert trees[0].get_children() and len(trees[0].get_children()) == 2

    win.destroy()
    app.destroy()
    db.close()
print('PRODUCTION_DASHBOARD_UI_SMOKETEST_OK')
