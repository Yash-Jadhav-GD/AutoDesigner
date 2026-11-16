"""
AutoDesigner Icon and Badge Service
Assigns semantic badges, numbers, or glyphs based on text keywords and layout context.
"""

from typing import Optional


class IconService:
    """Provides badge labels and semantic unicode glyphs for card items."""

    KEYWORD_GLYPHS = {
        "security": "🛡️",
        "lock": "🔒",
        "protect": "🛡️",
        "data": "📊",
        "analytics": "📈",
        "metric": "📈",
        "growth": "🚀",
        "performance": "⚡",
        "speed": "⚡",
        "cloud": "☁️",
        "cost": "💰",
        "financial": "💰",
        "dollar": "💲",
        "user": "👤",
        "team": "👥",
        "people": "👥",
        "target": "🎯",
        "goal": "🎯",
        "setting": "⚙️",
        "process": "🔄",
        "workflow": "🔄",
        "integration": "🔗",
        "ai": "🤖",
        "smart": "💡",
        "idea": "💡",
        "check": "✔️",
        "compliance": "📋",
        "document": "📄",
        "globe": "🌐",
        "global": "🌐",
        "mobile": "📱",
        "api": "🔌",
        "time": "⏱️",
        "schedule": "📅",
    }

    @classmethod
    def get_badge_text(cls, text: str, index: int, use_glyph: bool = False) -> str:
        """Returns either a numeric badge ('1', '2'...) or a semantic icon glyph."""
        if not use_glyph:
            return str(index + 1)

        lower = text.lower()
        for kw, glyph in cls.KEYWORD_GLYPHS.items():
            if kw in lower:
                return glyph

        return str(index + 1)
