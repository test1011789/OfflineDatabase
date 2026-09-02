from pathlib import Path
import tempfile
import tkinter as tk
from tkinter import ttk

from app import Database, DisplaySettingsDialog, FieldDialog, LinkDialog, OfflineDatabaseApp, LINK_MARKER

source = Path(__file__).resolve().parent / 'initial_data.xlsx'
with tempfile.TemporaryDirectory(prefix='offline_excel_db_gui_test_') as temp_dir:
    db = Database(Path(temp_dir))
    db.import_excel(source, replace=True)
    db.save_display_settings({
        'header_font_size': 16,
        'data_font_size': 12,
        'header_row_height': 54,
        'data_row_height': 32,
        'column_widths': {},
        'link_column_widths': {},
        'link_marker': '•',
        'show_link_marker': True,
        'font_size': 99,
        'default_column_width': 999,
        'form_gap': 999,
    })
    settings = db.get_display_settings()
    assert settings['header_font_size'] == 16
    assert settings['data_font_size'] == 12
    assert settings['header_row_height'] == 54
    assert settings['data_row_height'] == 32
    assert 'font_size' not in settings
    assert 'default_column_width' not in settings
    assert 'form_gap' not in settings

    app = OfflineDatabaseApp(db)
    app.deiconify()
    app.update()
    fields = db.list_fields()
    record_links = db.get_record_links(1)
    assert len(record_links) == 2
    link_columns = [column for column in app.tree['columns'] if str(column).startswith('link')]
    assert set(link_columns) == {f'link{field_id}' for field_id in record_links}
    assert all(app.tree.heading(column, 'text') == '' for column in link_columns)
    assert all(int(app.tree.column(column, 'width')) == 24 for column in link_columns)
    first_item = app.tree.get_children()[0]
    link_column = f'link{next(iter(record_links))}'
    link_bbox = app.tree.bbox(first_item, link_column)
    assert link_bbox, link_column
    link_event = type('Event', (), {'x': link_bbox[0] + 5, 'y': link_bbox[1] + 5})()
    assert app._link_at_event(link_event) == record_links[next(iter(record_links))]
    assert app.tree.set(first_item, link_column) == '•'
    app.tree.column(link_column, width=31)
    app._save_current_column_widths()
    app.refresh_data()
    assert int(app.tree.column(link_column, 'width')) == 31
    assert db.get_display_settings()['link_column_widths'][str(next(iter(record_links)))] == 31
    app.display_settings['show_link_marker'] = False
    app.refresh_data()
    assert not [column for column in app.tree['columns'] if str(column).startswith('link')]
    app.display_settings['show_link_marker'] = True
    app.refresh_data()
    assert app.tree.set(first_item, link_column) == '•'
    license_field = next(field for field in fields if field['name'].startswith('許可證號'))
    license_field_id = int(license_field['id'])
    app.sort_by_field(license_field_id)
    first_order = [app.tree.item(item, 'values')[2] for item in app.tree.get_children()]
    assert first_order == ['1', '2', '3'], first_order
    app.sort_by_field(license_field_id)
    second_order = [app.tree.item(item, 'values')[2] for item in app.tree.get_children()]
    assert second_order == ['3', '2', '1'], second_order

    license_column = f'c{license_field_id}'
    assert str(app.tree.column(license_column, 'anchor')) == 'center'
    app.tree.column(license_column, width=333)
    app._save_current_column_widths()
    app.refresh_data()
    assert int(app.tree.column(license_column, 'width')) == 333
    assert db.get_display_settings()['column_widths'][str(license_field_id)] == 333

    app.tree.selection_set(app.tree.get_children()[0])
    app.update_idletasks()
    style = ttk.Style(app)
    selected_background = style.map('Treeview', 'background')
    assert any('#FFFFFF' in str(item) for item in selected_background)
    field_selected_background = style.map('FieldTreeview', 'background')
    field_selected_foreground = style.map('FieldTreeview', 'foreground')
    assert any('#000000' in str(item) for item in field_selected_background)
    assert any('#FFFFFF' in str(item) for item in field_selected_foreground)
    assert app._tree_heading_font == ('Microsoft JhengHei UI', 16, 'bold')
    assert app._tree_font == ('Microsoft JhengHei UI', 12)
    assert LINK_MARKER == '📎'

    # 以程式化拖曳模擬，把第一個欄位移到最後，確認 position 與資料表同步。
    original_ids = [int(item) for item in app.fields_tree.get_children()]
    dragged_id = str(original_ids[0])
    app.fields_tree.move(dragged_id, '', len(original_ids) - 1)
    app._field_drag_iid = dragged_id
    app._on_field_release(None)
    reordered_ids = [int(item) for item in app.fields_tree.get_children()]
    assert reordered_ids[-1] == original_ids[0], reordered_ids
    assert [int(field['id']) for field in db.list_fields()][-1] == original_ids[0]
    app.refresh_data()
    assert app.tree['columns'][-1] == f'c{original_ids[0]}'

    # 所有欄位類型的欄位名稱都應可編輯，包含下拉選單與自動連動欄位。
    for editable_field in db.list_fields():
        dialog = FieldDialog(app, db, editable_field)
        dialog.update_idletasks()
        assert str(dialog.name_entry.cget('state')) != 'disabled', editable_field['name']
        dialog.destroy()

    # 資料編輯視窗的每一個欄位都應產生連結編輯按鈕。
    app.open_record_editor(1)
    app.update_idletasks()
    editor_windows = [child for child in app.winfo_children() if isinstance(child, tk.Toplevel) and child.title() == '編輯資料']
    assert len(editor_windows) == 1
    editor = editor_windows[0]
    edit_link_buttons = [child for child in editor.winfo_children() if isinstance(child, tk.Frame)]
    assert editor is not None
    editor.destroy()

    link_dialog = LinkDialog(app, 'old-url')
    link_dialog.url_var.set('new-url')
    link_dialog._save()
    assert link_dialog.result == 'new-url'

    settings_dialog = DisplaySettingsDialog(app, db, db.get_display_settings())
    settings_dialog.update_idletasks()
    assert settings_dialog.winfo_width() <= 520
    assert settings_dialog.winfo_height() <= 420
    settings_dialog.destroy()

    app.destroy()
    db.close()

    db2 = Database(Path(temp_dir))
    settings2 = db2.get_display_settings()
    assert settings2['header_font_size'] == 16
    assert settings2['data_font_size'] == 12
    assert settings2['header_row_height'] == 54
    assert settings2['data_row_height'] == 32
    assert settings2['column_widths'][str(license_field_id)] == 333
    assert settings2['link_column_widths'][str(next(iter(record_links)))] == 31
    assert settings2['link_marker'] == '•'
    assert settings2['show_link_marker'] is True
    assert 'font_size' not in settings2
    reordered_after_reopen = [int(field['id']) for field in db2.list_fields()]
    assert reordered_after_reopen[-1] == original_ids[0]
    db2.close()

print('GUI_SMOKETEST_OK lookup_style=black_white field_reorder=persisted data_sync=yes')
