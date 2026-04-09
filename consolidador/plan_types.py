from __future__ import annotations

PLAN_TYPES: tuple[str, str] = ("saude", "odonto")

PLAN_LABELS = {
    "saude": "Saude",
    "odonto": "Odonto",
}

OUTPUT_BASENAMES = {
    "saude": "consolidado_saude",
    "odonto": "consolidado_odonto",
}

MAIN_SHEET_NAMES = {
    "saude": "Consolidado Saude",
    "odonto": "Consolidado Odonto",
}


def is_valid_plan_type(plan_type: str | None) -> bool:
    return bool(plan_type) and plan_type in PLAN_TYPES


def plan_label(plan_type: str | None) -> str:
    if not plan_type:
        return "Nao informado"
    return PLAN_LABELS.get(plan_type, str(plan_type).title())
