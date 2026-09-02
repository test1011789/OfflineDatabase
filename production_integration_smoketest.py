from pathlib import Path
import tempfile
import tkinter as tk
from tkinter import ttk
from app import Database, OfflineDatabaseApp

source = Path(__file__).resolve().parent / 'initial_data.xlsx'
with tempfile.TemporaryDirectory(prefix='production_integration_') as temp_dir:
    db = Database(Path(temp_dir))
    db.import_excel(source, replace=True)
    app = OfflineDatabaseApp(db)
    app.update_idletasks()
    fields = db.list_fields()
    brand = next(f for f in fields if f['system_key'] == 'brand_name')
    production_cols = [c for c in app.tree['columns'] if str(c) == f"production{int(brand['id'])}"]
    assert production_cols == [f"production{int(brand['id'])}"], production_cols
    assert app.tree.heading(production_cols[0], 'text') == ''
    assert int(app.tree.column(production_cols[0], 'width')) == 34
    first_item = app.tree.get_children()[0]
    bbox = app.tree.bbox(first_item, production_cols[0])
    assert bbox, 'production icon cell bbox missing'
    event = type('Event', (), {'x': bbox[0] + 5, 'y': bbox[1] + 5})()
    assert app._production_record_at_event(event) == 1
    app.open_production_records(1)
    app.update_idletasks()
    wins = [child for child in app.winfo_children() if isinstance(child, tk.Toplevel) and child.title() == '生產記錄']
    assert len(wins) == 1, len(wins)
    win = wins[0]
    win.destroy()
    app.destroy()
    db.close()
print('PRODUCTION_UI_SMOKETEST_OK')
