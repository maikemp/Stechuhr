"""Configuration management for Stechuhr."""

import json
import os
from pathlib import Path

DEFAULT_CONFIG = {
    "data_dir": "",  # filled in at runtime with config_dir/data
    "travel_offset_minutes": 2,
    "expected_hours": {
        "monday": 8.0,
        "tuesday": 8.0,
        "wednesday": 8.0,
        "thursday": 8.0,
        "friday": 8.0,
    },
    "auto_break_threshold_hours": 6.0,
    "auto_break_deduction_minutes": 30,
    "carry_over_balance": {},
}

WEEKDAY_KEYS = ["monday", "tuesday", "wednesday", "thursday", "friday"]


def get_config_dir() -> Path:
    xdg = os.environ.get("XDG_CONFIG_HOME")
    if xdg:
        base = Path(xdg)
    else:
        base = Path.home() / ".config"
    d = base / "stechuhr"
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_config_path() -> Path:
    return get_config_dir() / "config.json"


def load_config() -> dict:
    path = get_config_path()
    if path.exists():
        with open(path) as f:
            cfg = json.load(f)
    else:
        cfg = {}

    # Merge with defaults so new keys are always present
    merged = {**DEFAULT_CONFIG, **cfg}

    # Ensure data_dir has a value
    if not merged["data_dir"]:
        merged["data_dir"] = str(get_config_dir() / "data")

    # Ensure carry_over_balance exists
    if "carry_over_balance" not in merged:
        merged["carry_over_balance"] = {}

    # Save if newly created
    if not path.exists():
        save_config(merged)

    return merged


def save_config(cfg: dict) -> None:
    path = get_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def get_data_dir(cfg: dict) -> Path:
    d = Path(os.path.expanduser(cfg["data_dir"]))
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_expected_hours(cfg: dict, weekday: int) -> float:
    """Return expected hours for a weekday (0=Monday, 4=Friday)."""
    if weekday < 0 or weekday > 4:
        return 0.0
    key = WEEKDAY_KEYS[weekday]
    return cfg["expected_hours"].get(key, 0.0)


def get_travel_offset(cfg: dict) -> int:
    return cfg.get("travel_offset_minutes", 2)


def get_carry_over(cfg: dict, year: int) -> float:
    return cfg.get("carry_over_balance", {}).get(str(year), 0.0)


def set_carry_over(cfg: dict, year: int, balance: float) -> None:
    cfg.setdefault("carry_over_balance", {})[str(year)] = round(balance, 4)
    save_config(cfg)
