from __future__ import annotations

import os
import re


RE_COLOR_PRINT = (
    "大身：净色\n"
    "内衬：配色\n\n"
    "注：\n"
    "1.定位问题：——\n"
    "2. 印花朝向问题：——\n"
    "3.条纹/格子款式是否对条/对格：——"
)

RE_WORKMANSHIP_BY_ORDER = {
    "805CM": "BRLPS428",
    "805CN": "BRLPS429",
    "803CN": "BRLPS430",
    "804CM": "BRLPS431",
    "804CN": "BRLPS432",
}


def _clean(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _order_from_text_or_filename(text: str, filename: str = "", current_order_id: str = "") -> str:
    for value in (current_order_id, filename, text):
        match = re.search(r"\b\d{3}C[MN]\b", str(value or ""), flags=re.I)
        if match:
            return match.group(0).upper()
    return _clean(current_order_id)


def normalize_re_product_name(text: str, order_id: str = "") -> str:
    source = f"{order_id}\n{text or ''}".lower()
    if "op01" in source or "one piece" in source or "803cn" in source:
        return "女士连体泳装"
    if order_id.endswith("CM") or "soft bra" in source or "top05" in source or "removable cup" in source:
        return "女士塞杯泳装上衣"
    if "bk51" in source or "high waisted" in source:
        return "女士高腰泳裤"
    if "bk01" in source or "basic classical with tie" in source:
        return "女士系带泳裤"
    if order_id.endswith("CN"):
        return "女士泳裤"
    return ""


def format_re_fabric_quality(order_id: str, text: str = "", raw_fabric_quality: str = "") -> str:
    order_id = _clean(order_id).upper()
    if order_id == "803CN":
        return (
            "大身：85%尼龙，15%氨纶，200G\n\n"
            "前后片内衬：90%涤纶，10%氨纶，130G\n\n"
            "塞杯牛奶丝：100%涤纶，同上一季操作RE订单"
        )
    if order_id.endswith("CM"):
        return (
            "大身：85%尼龙，15%氨纶\n\n"
            "塞杯牛奶丝：100%涤纶，同上一季操作RE订单"
        )
    if order_id.endswith("CN"):
        return (
            "大身：85%尼龙，15%氨纶\n\n"
            "前后片内衬：90%涤纶，10%氨纶，130G"
        )
    return _clean(raw_fabric_quality)


def apply_re_quote_rules(text: str, data: dict, filename: str = "") -> dict:
    result = dict(data)
    order_id = _order_from_text_or_filename(text, filename, result.get("order_id", ""))
    if order_id:
        result["order_id"] = order_id

    product_name = normalize_re_product_name(text, order_id)
    if product_name:
        result["product_name"] = product_name

    fabric_quality = format_re_fabric_quality(order_id, text, result.get("fabric_quality", ""))
    if fabric_quality:
        result["fabric_quality"] = fabric_quality

    result["color_print"] = RE_COLOR_PRINT
    if order_id in RE_WORKMANSHIP_BY_ORDER:
        result["workmanship_image"] = RE_WORKMANSHIP_BY_ORDER[order_id]
    return result
