"""Color mapping utilities for DOCX plugin.

This module contains Word-specific color processing logic that was moved from
the core Orlando Toolkit to the DOCX plugin during the plugin architecture refactor.
It handles Word theme colors, highlight backgrounds, and HSV-based color detection.
"""

from typing import Optional, Dict, Any


def convert_color_to_outputclass(
    color_value: Optional[str], color_rules: Dict[str, Any]
) -> Optional[str]:
    """Map Word colour representation to Orlando `outputclass` (red/green).

    The logic is unchanged from the legacy implementation, supporting:
    • exact hex matches (case-insensitive)
    • a limited set of Word theme colours
    • heuristic detection based on RGB dominance.
    
    Args:
        color_value: Color value from Word document (hex, theme-, background- prefixed)
        color_rules: Color mapping rules from plugin configuration
        
    Returns:
        Orlando outputclass string (e.g., 'color-red', 'color-green') or None
    """
    if not color_value:
        return None

    color_mappings = color_rules.get("color_mappings", {})
    theme_map = color_rules.get("theme_map", {})

    color_lower = color_value.lower()
    if color_lower in color_mappings:
        return color_mappings[color_lower]

    if color_value.startswith("theme-"):
        theme_name = color_value[6:]
        return theme_map.get(theme_name)

    # Background colour tokens coming from shading (already prefixed)
    if color_value.startswith("background-"):
        return color_mappings.get(color_value)

    # ------------------------------------------------------------------
    # HSV-based tolerance fallback (optional)
    # ------------------------------------------------------------------
    tolerance_cfg = color_rules.get("tolerance", {})
    if color_lower.startswith("#") and len(color_lower) == 7 and tolerance_cfg:
        try:
            r = int(color_lower[1:3], 16) / 255.0
            g = int(color_lower[3:5], 16) / 255.0
            b = int(color_lower[5:7], 16) / 255.0

            import colorsys

            h, s, v = colorsys.rgb_to_hsv(r, g, b)
            h_deg = h * 360
            s_pct = s * 100
            v_pct = v * 100

            for out_class, spec in tolerance_cfg.items():
                # Extract ranges
                hue_range = spec.get("hue")
                hue2_range = spec.get("hue2")  # optional secondary segment (wrap-around)
                sat_min = spec.get("sat_min", 0)
                val_min = spec.get("val_min", 0)

                def _in_range(hrange: list[int] | tuple[int, int] | None) -> bool:
                    if not hrange:
                        return False
                    start, end = hrange
                    return start <= h_deg <= end

                if (
                    (_in_range(hue_range) or _in_range(hue2_range))
                    and s_pct >= sat_min
                    and v_pct >= val_min
                ):
                    return out_class
        except Exception:
            pass

    return None