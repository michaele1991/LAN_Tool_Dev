from smbus_parser_tool.parser import records_from_rows, summarize


def test_records_from_rows_accepts_common_columns():
    rows = [{"Time": "0.1", "Address": "0x50", "Operation": "Write", "Data": "01 02"}]
    records = list(records_from_rows(rows))
    assert records[0].timestamp == "0.1"
    assert records[0].address == "0x50"
    assert records[0].operation == "Write"
    assert "Records: 1" in summarize(records)
