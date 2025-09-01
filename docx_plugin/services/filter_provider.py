from __future__ import annotations

from typing import Dict, List, Optional, Set

from orlando_toolkit.core.plugins.interfaces import FilterProvider
from orlando_toolkit.core.models import DitaContext
from orlando_toolkit.core.services import heading_analysis_service as _heading_analysis


class DocxFilterProvider(FilterProvider):
    """FilterProvider implementation for DOCX sources.

    Delegates heading/style analysis to core heading_analysis_service and
    maps plugin panel exclusions to the structure editing service expectations.
    """

    def get_counts(self, context: DitaContext) -> Dict[str, int]:
        try:
            original_root = getattr(context, 'ditamap_root', None)
            if hasattr(context, 'restore_from_original'):
                context.restore_from_original()
            result = _heading_analysis.build_headings_cache(context)
            if original_root is not None:
                context.ditamap_root = original_root
            return result
        except Exception:
            return {}

    def get_occurrences(self, context: DitaContext) -> Dict[str, List[Dict[str, str]]]:
        try:
            original_root = getattr(context, 'ditamap_root', None)
            if hasattr(context, 'restore_from_original'):
                context.restore_from_original()
            result = _heading_analysis.build_heading_occurrences(context)
            if original_root is not None:
                context.ditamap_root = original_root
            return result
        except Exception:
            return {}

    def get_occurrences_current(self, context: DitaContext) -> Dict[str, List[Dict[str, str]]]:
        try:
            return _heading_analysis.build_heading_occurrences(context)
        except Exception:
            return {}

    def get_levels(self, context: DitaContext) -> Dict[str, Optional[int]]:
        try:
            original_root = getattr(context, 'ditamap_root', None)
            if hasattr(context, 'restore_from_original'):
                context.restore_from_original()
            result = _heading_analysis.build_style_levels(context)
            if original_root is not None:
                context.ditamap_root = original_root
            return result
        except Exception:
            return {}

    def build_exclusion_map(self, exclusions: Dict[str, bool]) -> Dict[int, Set[str]]:
        # Map excluded style flags to per-level set using current (original) levels
        try:
            # Provider does not receive context here by design; callers compute levels first when needed
            # For simplicity, assume UI will call get_levels via controller and then call this method
            # with styles aligned to known levels. Fallback to level 1 when unknown.
            # Since we don't have levels here, this function is best-effort and expects
            # the controller to compute levels via get_levels and then call provider again.
            # To keep contract simple, do a level-1 mapping only.
            style_excl_map: Dict[int, Set[str]] = {}
            for style, excluded in (exclusions or {}).items():
                if not excluded:
                    continue
                style_excl_map.setdefault(1, set()).add(style)
            return style_excl_map
        except Exception:
            return {}

    def estimate_unmergable(self, context: DitaContext, style_excl_map: Dict[int, Set[str]]) -> int:
        try:
            return int(_heading_analysis.count_unmergable_for_styles(context, style_excl_map))
        except Exception:
            return 0

