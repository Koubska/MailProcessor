"""What the GUI needs from the system: sharp display on scaled Windows screens (125 %, 150 %), opening files."""

import sys
from pathlib import Path

import pytest

import mailprocessor.gui.desktop as gui_module
from mailprocessor.gui.desktop import enable_windows_dpi_awareness, open_in_default_app, ui_scale


class _Untouchable:
    def __getattr__(self, name: str):
        raise AssertionError(f"Windows API used on another platform: {name}")


class _Api:
    def __init__(self, fail: bool = False) -> None:
        self.calls: list[tuple] = []
        self.fail = fail

    def __call__(self, *args):
        if self.fail:
            raise OSError("not available")
        self.calls.append(args)
        return 0


class _Windll:
    def __init__(self, shcore=None, user32=None) -> None:
        if shcore is not None:
            self.shcore = shcore
        if user32 is not None:
            self.user32 = user32


class _Library:
    def __init__(self, **functions) -> None:
        self.__dict__.update(functions)


@pytest.mark.parametrize("platform", ["darwin", "linux"])
def test_nothing_happens_outside_windows(platform: str) -> None:
    assert enable_windows_dpi_awareness(platform, windll=_Untouchable()) is False


def test_windows_8_1_and_later_use_system_dpi_awareness() -> None:
    set_awareness = _Api()

    assert enable_windows_dpi_awareness("win32", windll=_Windll(shcore=_Library(SetProcessDpiAwareness=set_awareness)))
    # 1 = system DPI aware: Tk 8.6 cannot re-layout when a window moves between monitors with different scaling.
    assert set_awareness.calls == [(1,)]


def test_older_windows_fall_back_to_user32() -> None:
    set_aware = _Api()

    assert enable_windows_dpi_awareness("win32", windll=_Windll(user32=_Library(SetProcessDPIAware=set_aware)))
    assert set_aware.calls == [()]


def test_failing_calls_never_stop_the_app() -> None:
    windll = _Windll(
        shcore=_Library(SetProcessDpiAwareness=_Api(fail=True)), user32=_Library(SetProcessDPIAware=_Api(fail=True))
    )

    assert enable_windows_dpi_awareness("win32", windll=windll) is False


@pytest.mark.parametrize(("pixels_per_inch", "scale"), [(96, 1.0), (120, 1.25), (144, 1.5), (192, 2.0), (72, 1.0)])
def test_ui_scale_follows_the_screen_but_never_shrinks(pixels_per_inch: float, scale: float) -> None:
    assert ui_scale(pixels_per_inch) == scale


@pytest.mark.skipif(sys.platform != "win32", reason="real Windows API, runs on the Windows CI runner")
def test_real_windows_api_and_tk_window() -> None:
    import tkinter

    assert enable_windows_dpi_awareness(sys.platform) is True
    try:
        root = tkinter.Tk()
    except tkinter.TclError as exc:
        pytest.skip(f"no display: {exc}")
    try:
        assert ui_scale(root.winfo_fpixels("1i")) >= 1.0
    finally:
        root.destroy()


@pytest.mark.parametrize(("platform", "command"), [("darwin", "open"), ("linux", "xdg-open")])
def test_open_in_default_app_uses_the_system_opener(monkeypatch, tmp_path: Path, platform: str, command: str) -> None:
    calls = []
    monkeypatch.setattr(gui_module.sys, "platform", platform)
    monkeypatch.setattr(gui_module.subprocess, "Popen", lambda args: calls.append(args))

    open_in_default_app(tmp_path / "out.xlsx")

    assert calls == [[command, str(tmp_path / "out.xlsx")]]
