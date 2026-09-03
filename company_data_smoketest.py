from __future__ import annotations

import tempfile
from pathlib import Path

from app import Database


def main() -> None:
    root = Path(tempfile.mkdtemp(prefix='offline_db_company_test_'))
    db = Database(root)
    companies = db.companies()
    assert len(companies) >= 2

    company_a = int(companies[0]['id'])
    company_b = int(companies[1]['id'])

    field_id = db.add_field('公司測試欄位', 'text', [])
    record_a = db.create_record({field_id: 'A'}, company_id=company_a)
    record_b = db.create_record({field_id: 'B'}, company_id=company_b)

    assert [r['id'] for r in db.all_records(company_a)] == [record_a]
    assert [r['id'] for r in db.all_records(company_b)] == [record_b]

    db.ensure_production_manufacturer('同名廠商', company_a)
    db.ensure_production_manufacturer('同名廠商', company_b)
    assert len(db.production_manufacturers(company_a)) == 1
    assert len(db.production_manufacturers(company_b)) == 1

    db.create_production_record(record_a, '2026-09-03', '100', '90', '同名廠商')
    db.create_production_record(record_b, '2026-09-03', '100', '100', '同名廠商')
    assert db.production_record_count(record_a, abnormal_only=True) == 1
    assert db.production_record_count(record_b, abnormal_only=True) == 0

    db.close()
    print('V7_COMPANY_TEST_OK')


if __name__ == '__main__':
    main()
