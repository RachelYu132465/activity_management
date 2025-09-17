from scripts.core.data_util import load_program_by_id
from scripts.actions.add_program_data_to_excel import _collect_program_columns


def test_load_program_by_id_returns_exact_match():
    program = load_program_by_id(2, fallback_to_first=False)
    assert str(program.get("id")) == "2"


def test_collect_program_columns_handles_missing_values():
    program = {
        "eventNames": ["Main Event"],
        "date": "2025-01-02",
        "locations": ["Location A"],
    }
    columns = dict(_collect_program_columns(program))
    assert columns["program_data.eventNames[0]"] == "Main Event"
    assert columns["program_data.eventNames[1]"] == ""
    assert columns["program_data.date"] == "2025-01-02"
    assert columns["program_data.locations[0]"] == "Location A"
    assert columns["program_data.locations[1]"] == ""
