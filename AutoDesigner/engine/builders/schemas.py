"""
Data schemas for every registered archetype.
Used to validate specs before building, and to tell Claude the exact shape to fill.
Shapes match what the builders in archetypes.py / generic.py read.
"""

from typing import Dict, List, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field, TypeAdapter


class _Strict(BaseModel):
    model_config = ConfigDict(extra="forbid", populate_by_name=True)


# --- Branded archetypes (engine/builders/archetypes.py) ----------------------

class Column(_Strict):
    header: str
    items: List[str]


class Pair(_Strict):
    subheader: str
    body: str


class SplitProcessInsights(_Strict):
    left_pairs: List[Pair]
    right_header: str
    right_items: List[str]


class KpiRow(_Strict):
    val: str
    label: str
    sub: str


class AssessmentKpis(_Strict):
    title: str
    subtitle: str
    left_pairs: List[Pair]
    deliv_header: str
    deliv_items: List[str]
    insights_header: str
    kpi_rows: List[KpiRow]


class Workflow(_Strict):
    letter: str
    title: str
    desc: str
    in_: str = Field(alias="in")
    out: str
    autonomy: str


class WorkflowCards(_Strict):
    title: str
    subtitle: str
    workflows: List[Workflow]


class PersonaCard(_Strict):
    count: str
    label: str
    name: str
    regions: str


class PersonaGrid(_Strict):
    title: str
    subtitle: str
    lead_badge: str
    total_badge: str
    persona_cards: List[PersonaCard]


class Leader(_Strict):
    title: str
    sub: Optional[str]
    ftes: str
    type: Literal["spend_value", "span_breaker"]


class CardPool(_Strict):
    title: str
    badge: str
    sub: str
    cards: List[str]


class RegionPool(_Strict):
    title: str
    badge: str
    sub: str
    regions: List[str]


class Legend(_Strict):
    l1: str
    l2: str


class OrgStructure(_Strict):
    title: str
    top_banner: str
    leadership: List[Leader]
    pool1: CardPool
    pool2: RegionPool
    pool3: CardPool
    legend: Legend


class AllocationTable(_Strict):
    title: Optional[str]
    note: Optional[str]
    footer: Optional[str]
    table_data: List[List[List[str]]]


class SizingTable(_Strict):
    title: Optional[str]
    table_data: List[List[List[str]]]


# --- Generic card archetypes (engine/builders/generic.py) --------------------

class Item(_Strict):
    subheader: str = ""
    body: str = ""


class Card(_Strict):
    header: str
    items: List[Item] = []
    kpi_metric: Optional[str] = None
    kpi_label: Optional[str] = None


class CardSlide(_Strict):
    title: str = ""
    subtitle: str = ""
    cards: List[Card]


DATA_TYPES: Dict[str, object] = {
    "three_column_pills": List[Column],
    "split_process_insights": SplitProcessInsights,
    "process_rows": List[Pair],
    "assessment_kpis": AssessmentKpis,
    "workflow_cards": WorkflowCards,
    "persona_grid": PersonaGrid,
    "org_structure": OrgStructure,
    "allocation_table": AllocationTable,
    "sizing_table": SizingTable,
    "card_grid": CardSlide,
    "kpi_cards": CardSlide,
    "subheader_list": CardSlide,
    "data_box": CardSlide,
    "crd_icn_v1": CardSlide,
    "important_box": CardSlide,
    "imp_box_v1": CardSlide,
}

ADAPTERS: Dict[str, TypeAdapter] = {name: TypeAdapter(t) for name, t in DATA_TYPES.items()}


def validate_data(archetype: str, data):
    """Validates and normalises slide data; returns plain JSON-able data (aliases kept)."""
    adapter = ADAPTERS[archetype]
    return adapter.dump_python(adapter.validate_python(data), by_alias=True, mode="json")


def json_schema(archetype: str) -> dict:
    return ADAPTERS[archetype].json_schema(by_alias=True)
