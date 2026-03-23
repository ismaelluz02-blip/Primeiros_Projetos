from __future__ import annotations

import ctypes
import time
from ctypes import wintypes
from typing import Any


user32 = ctypes.windll.user32

MONITORINFOF_PRIMARY = 1
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
KEYEVENTF_KEYUP = 0x0002

VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_BACK = 0x08
VK_F2 = 0x71


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", wintypes.LONG),
        ("top", wintypes.LONG),
        ("right", wintypes.LONG),
        ("bottom", wintypes.LONG),
    ]


class MONITORINFOEXW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("rcMonitor", RECT),
        ("rcWork", RECT),
        ("dwFlags", wintypes.DWORD),
        ("szDevice", wintypes.WCHAR * 32),
    ]


def set_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass

    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass


def get_mouse_position() -> tuple[int, int]:
    point = POINT()
    user32.GetCursorPos(ctypes.byref(point))
    return int(point.x), int(point.y)


def set_mouse_position(x: int, y: int) -> None:
    user32.SetCursorPos(int(x), int(y))


def left_click(clicks: int = 1, interval: float = 0.08) -> None:
    repeat = max(1, int(clicks))
    for idx in range(repeat):
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        if idx < repeat - 1:
            time.sleep(max(0.0, interval))


def _key_down(vk_code: int) -> None:
    user32.keybd_event(int(vk_code), 0, 0, 0)


def _key_up(vk_code: int) -> None:
    user32.keybd_event(int(vk_code), 0, KEYEVENTF_KEYUP, 0)


def press_key(vk_code: int, hold: float = 0.02) -> None:
    _key_down(vk_code)
    time.sleep(max(0.0, hold))
    _key_up(vk_code)


def press_f2() -> None:
    press_key(VK_F2)


def press_backspace(times: int = 1, interval: float = 0.03) -> None:
    repeat = max(1, int(times))
    for idx in range(repeat):
        press_key(VK_BACK, hold=0.01)
        if idx < repeat - 1:
            time.sleep(max(0.0, interval))


def type_text(text: str, key_delay: float = 0.015) -> None:
    for char in text:
        vk_scan = user32.VkKeyScanW(ord(char))
        if vk_scan == -1:
            continue

        vk_code = vk_scan & 0xFF
        shift_state = (vk_scan >> 8) & 0xFF

        mods: list[int] = []
        if shift_state & 1:
            mods.append(VK_SHIFT)
        if shift_state & 2:
            mods.append(VK_CONTROL)
        if shift_state & 4:
            mods.append(VK_MENU)

        for mod in mods:
            _key_down(mod)

        press_key(vk_code, hold=0.01)

        for mod in reversed(mods):
            _key_up(mod)

        time.sleep(max(0.0, key_delay))


def list_monitors() -> list[dict[str, Any]]:
    monitors: list[dict[str, Any]] = []

    monitor_enum_proc = ctypes.WINFUNCTYPE(
        wintypes.BOOL,
        wintypes.HMONITOR,
        wintypes.HDC,
        ctypes.POINTER(RECT),
        wintypes.LPARAM,
    )

    @monitor_enum_proc
    def callback(h_monitor, _hdc, _rect, _data) -> bool:
        info = MONITORINFOEXW()
        info.cbSize = ctypes.sizeof(MONITORINFOEXW)
        user32.GetMonitorInfoW(h_monitor, ctypes.byref(info))

        left = int(info.rcMonitor.left)
        top = int(info.rcMonitor.top)
        right = int(info.rcMonitor.right)
        bottom = int(info.rcMonitor.bottom)
        width = right - left
        height = bottom - top

        monitors.append(
            {
                "x": left,
                "y": top,
                "width": width,
                "height": height,
                "is_primary": bool(info.dwFlags & MONITORINFOF_PRIMARY),
                "device": info.szDevice,
            }
        )
        return True

    user32.EnumDisplayMonitors(0, 0, callback, 0)
    monitors.sort(key=lambda item: (item["x"], item["y"]))
    return monitors
