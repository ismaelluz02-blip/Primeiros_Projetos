from __future__ import annotations

from pathlib import Path
from typing import Any

try:
    import cv2  # type: ignore
    import numpy as np  # type: ignore
    from mss import mss  # type: ignore
except Exception:  # pragma: no cover - optional dependency
    cv2 = None
    np = None
    mss = None


def is_vision_available() -> bool:
    return bool(cv2 is not None and np is not None and mss is not None)


def missing_vision_message() -> str:
    return (
        "Modo por imagem indisponivel. Instale dependencias: "
        "pip install -r requirements.txt"
    )


def _grab_region(left: int, top: int, width: int, height: int):
    if not is_vision_available():
        raise RuntimeError(missing_vision_message())

    safe_w = max(1, int(width))
    safe_h = max(1, int(height))
    with mss() as screen:
        shot = screen.grab({"left": int(left), "top": int(top), "width": safe_w, "height": safe_h})
    img = np.array(shot)
    return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)


def save_template_around_point(
    *,
    x: int,
    y: int,
    output_path: Path,
    capture_size: int = 120,
    bounds: dict[str, int] | None = None,
) -> Path:
    if not is_vision_available():
        raise RuntimeError(missing_vision_message())

    size = max(40, int(capture_size))
    half = size // 2
    left = int(x) - half
    top = int(y) - half

    if bounds:
        min_x = int(bounds["x"])
        min_y = int(bounds["y"])
        max_x = min_x + int(bounds["width"]) - size
        max_y = min_y + int(bounds["height"]) - size
        left = min(max(left, min_x), max_x)
        top = min(max(top, min_y), max_y)

    image = _grab_region(left=left, top=top, width=size, height=size)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    ok = cv2.imwrite(str(output_path), image)
    if not ok:
        raise RuntimeError(f"Falha ao salvar template em {output_path}")
    return output_path


def locate_template_center(
    *,
    template_path: Path,
    monitor: dict[str, Any],
    threshold: float = 0.84,
) -> tuple[int, int, float] | None:
    if not is_vision_available():
        raise RuntimeError(missing_vision_message())
    if not template_path.exists():
        return None

    template = cv2.imread(str(template_path), cv2.IMREAD_COLOR)
    if template is None:
        return None

    width = int(monitor["width"])
    height = int(monitor["height"])
    screen = _grab_region(
        left=int(monitor["x"]),
        top=int(monitor["y"]),
        width=width,
        height=height,
    )

    t_h, t_w = template.shape[:2]
    s_h, s_w = screen.shape[:2]
    if t_h > s_h or t_w > s_w:
        return None

    result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
    _min_v, max_v, _min_loc, max_loc = cv2.minMaxLoc(result)
    score = float(max_v)
    if score < float(threshold):
        return None

    center_x = int(monitor["x"]) + int(max_loc[0]) + (t_w // 2)
    center_y = int(monitor["y"]) + int(max_loc[1]) + (t_h // 2)
    return center_x, center_y, score

