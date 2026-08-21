from __future__ import annotations

import json
import os
from functools import lru_cache

from excel_writer import SAMPLE_QUANTITY_NOTE


CONFIG_PATH = "customer_templates.json"
DEFAULT_TEMPLATE_CONFIG = {
    "customer_name": "默认模板",
    "template_path": "templates/报价单总表模板.xlsx",
    "sample_quantity_mode": "blank",
}


@lru_cache(maxsize=1)
def load_customer_template_configs() -> dict:
    if not os.path.exists(CONFIG_PATH):
        return {"default": DEFAULT_TEMPLATE_CONFIG.copy()}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    if "default" not in data:
        data["default"] = DEFAULT_TEMPLATE_CONFIG.copy()
    return data


def get_customer_template_config(customer_id: str | None) -> dict:
    configs = load_customer_template_configs()
    customer_id = str(customer_id or "").strip()
    config = configs.get(customer_id) or configs.get("default") or DEFAULT_TEMPLATE_CONFIG
    return {**DEFAULT_TEMPLATE_CONFIG, **config}


def get_sample_quantity_default(customer_id: str | None) -> str:
    config = get_customer_template_config(customer_id)
    if config.get("sample_quantity_mode") == "nkd_default":
        return SAMPLE_QUANTITY_NOTE
    return ""


def available_template_options(customer_id: str | None) -> list[dict]:
    config = get_customer_template_config(customer_id)
    templates = config.get("templates")
    if isinstance(templates, list) and templates:
        return [
            {
                "template_id": option.get("template_id") or f"{customer_id or 'default'}-{idx + 1}",
                "customer_name": option.get("customer_name") or config.get("customer_name") or str(customer_id or "默认模板"),
                "template_path": option.get("template_path") or config.get("template_path") or DEFAULT_TEMPLATE_CONFIG["template_path"],
                "sample_quantity_default": SAMPLE_QUANTITY_NOTE
                if option.get("sample_quantity_mode", config.get("sample_quantity_mode")) == "nkd_default"
                else "",
            }
            for idx, option in enumerate(templates)
        ]
    return [{
        "template_id": str(customer_id or "default"),
        "customer_name": config.get("customer_name") or str(customer_id or "默认模板"),
        "template_path": config.get("template_path") or DEFAULT_TEMPLATE_CONFIG["template_path"],
        "sample_quantity_default": get_sample_quantity_default(customer_id),
    }]
