from modules.data_loader import load_task_table


def test_load_csv_and_fill_defaults(tmp_path):
    data = tmp_path / "names.csv"
    data.write_text("task_id,name,status\n001,\u5f20\u4e09,\n", encoding="utf-8")
    table = load_task_table({"file": str(data), "key_column": "name"})
    assert table.rows[0]["keyword"] == "\u5f20\u4e09"
    assert table.rows[0]["status"] == "\u5f85\u5904\u7406"


def test_save_result_xlsx(tmp_path):
    data = tmp_path / "names.txt"
    data.write_text("\u5f20\u4e09\n\u674e\u56db\n", encoding="utf-8")
    table = load_task_table({"file": str(data)})
    result = table.save_xlsx(tmp_path / "result.xlsx")
    assert result.exists()
