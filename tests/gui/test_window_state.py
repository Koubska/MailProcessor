from pathlib import Path

import pytest

from mailprocessor.gui.window_state import WindowState, fit_geometry, load_window_state, save_window_state


def test_missing_or_broken_state_file_gives_defaults(tmp_path: Path) -> None:
    assert load_window_state(tmp_path / "missing.json") == WindowState()
    broken = tmp_path / "window.json"
    broken.write_text("{not json", encoding="utf-8")
    assert load_window_state(broken) == WindowState()
    broken.write_text('{"geometry": 5, "tab": "x"}', encoding="utf-8")
    assert load_window_state(broken) == WindowState()


def test_state_roundtrip_creates_the_folder(tmp_path: Path) -> None:
    path = tmp_path / "data" / "window.json"
    state = WindowState(geometry="1200x800+40+30", tab=2)

    save_window_state(path, state)

    assert load_window_state(path) == state


def test_saving_never_fails(tmp_path: Path) -> None:
    blocker = tmp_path / "data"
    blocker.write_text("a file where the folder should be", encoding="utf-8")

    save_window_state(blocker / "window.json", WindowState(tab=1))  # no exception


@pytest.mark.parametrize(
    ("saved", "expected"),
    [
        ("1200x800+40+30", "1200x800+40+30"),
        ("1200x800+-8+0", "1200x800+-8+0"),  # Windows reports small negative offsets for snapped windows
        ("1200x800+3000+30", None),  # was on a second monitor that is gone
        ("1200x800+40+1500", None),
        ("3000x800+0+0", None),  # larger than the screen
        ("garbage", None),
        (None, None),
    ],
)
def test_fit_geometry_only_restores_windows_that_are_visible(saved: str | None, expected: str | None) -> None:
    assert fit_geometry(saved, screen_width=1920, screen_height=1080) == expected
