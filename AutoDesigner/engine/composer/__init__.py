"""
Composable Layout Engine package.
Provides atomic layout primitives (VStack, HStack, Grid), a constraint-based box-model solver,
a semantic slide composer, and a universal DrawingML renderer enforcing 100% middle alignment,
exact 4.0pt corner radius, native DrawingML bullets, and 0 shape overlaps.
"""

from .primitives import (
    LayoutNode,
    HStack,
    VStack,
    Grid,
    CardNode,
    PillHeaderNode,
    BulletListNode,
    SectionListNode,
    SectionItem,
    KPINode,
    ProcessRowNode,
    TableNode,
    DividerNode,
    ConnectorNode,
    MaturityStageNode,
    SlideComposition
)
from .layout_solver import LayoutSolver, ResolvedShape
from .renderer import ComposableRenderer
from .composer import SemanticComposer

__all__ = [
    "LayoutNode",
    "HStack",
    "VStack",
    "Grid",
    "CardNode",
    "PillHeaderNode",
    "BulletListNode",
    "SectionListNode",
    "SectionItem",
    "KPINode",
    "ProcessRowNode",
    "TableNode",
    "DividerNode",
    "ConnectorNode",
    "MaturityStageNode",
    "SlideComposition",
    "LayoutSolver",
    "ResolvedShape",
    "ComposableRenderer",
    "SemanticComposer",
]
