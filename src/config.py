"""Load configuration from etc/config.yaml."""

from pathlib import Path
from typing import Any, Dict

import yaml


def load_config() -> Dict[str, Any]:
    """Load config from etc/config.yaml relative to project root."""
    # Find project root (parent of src/)
    src_dir = Path(__file__).parent
    project_root = src_dir.parent
    config_path = project_root / "etc" / "config.yaml"

    if not config_path.exists():
        # Return defaults if no config file
        return {
            "high_value": {
                "lifetime_revenue_min": 1_000_000,
                "peak_ttm_share_min": 0.02,
            },
            "status": {
                "active_max": 6,
                "inactive_max": 18,
            },
        }

    with open(config_path) as f:
        return yaml.safe_load(f)


# Convenience accessors
def get_high_value_thresholds() -> tuple:
    """Return (lifetime_revenue_min, peak_ttm_share_min)."""
    cfg = load_config()
    hv = cfg.get("high_value", {})
    return (
        hv.get("lifetime_revenue_min", 1_000_000),
        hv.get("peak_ttm_share_min", 0.02),
    )


def get_status_thresholds() -> tuple:
    """Return (active_max, inactive_max)."""
    cfg = load_config()
    st = cfg.get("status", {})
    return (
        st.get("active_max", 6),
        st.get("inactive_max", 18),
    )
