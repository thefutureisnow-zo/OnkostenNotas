"""
Tests voor state.py
"""
import pytest

from state import (
    load_state,
    save_state,
    is_processed,
    is_skipped,
    mark_processed,
    mark_skipped_weekend,
)


@pytest.fixture
def state_file(tmp_path):
    return tmp_path / "test_processed.json"


def test_empty_state(state_file):
    state = load_state(state_file)
    assert state == {"processed": [], "skipped_weekend": [], "metadata": {}}


def test_mark_and_check_processed(state_file):
    state = load_state(state_file)
    mark_processed("ABC123", state)
    assert is_processed("ABC123", state)
    assert not is_processed("XYZ999", state)


def test_mark_and_check_skipped(state_file):
    state = load_state(state_file)
    mark_skipped_weekend("WKD001", state)
    assert is_skipped("WKD001", state)
    assert not is_skipped("ABC123", state)


def test_persist_across_save_load(state_file):
    state = load_state(state_file)
    mark_processed("ABC123", state)
    mark_skipped_weekend("WKD001", state)
    save_state(state, state_file)

    reloaded = load_state(state_file)
    assert is_processed("ABC123", reloaded)
    assert is_skipped("WKD001", reloaded)


def test_no_duplicates(state_file):
    state = load_state(state_file)
    mark_processed("ABC123", state)
    mark_processed("ABC123", state)
    assert state["processed"].count("ABC123") == 1


def test_missing_file_returns_empty(tmp_path):
    state = load_state(tmp_path / "nonexistent.json")
    assert state["processed"] == []
