from pathlib import Path
import tempfile, sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from app import Database
OLD = Path('/mnt/data/app_20260907_093503.db')
XLSX = Path('/mnt/data/v10src/initial_data.xlsx')
with tempfile.TemporaryDirectory() as td:
    db = Database(Path(td))
    db.import_excel(XLSX, replace=True)
    company_ids=[int(r['id']) for r in db.companies()]
    info=db.inspect_legacy_data_db(OLD)
    assert info == {'record_count':49,'value_count':735,'link_count':88,'production_count':1}, info
    result=db.import_legacy_data(OLD, company_ids[0])
    assert result['records']==49, result
    assert result['values']==735 and result['links']==88 and result['production']==1, result
    assert result['unmapped_values']==0, result
    assert db.conn.execute('select count(*) from records where company_id=?',(company_ids[0],)).fetchone()[0]==52
    assert db.conn.execute('select count(*) from record_values rv join records r on r.id=rv.record_id where r.company_id=?',(company_ids[0],)).fetchone()[0]>=735
    assert db.conn.execute('select count(*) from records where company_id=?',(company_ids[1],)).fetchone()[0]==0
    print('legacy import smoketest passed', info, result)
