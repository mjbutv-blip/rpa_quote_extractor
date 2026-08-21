from __future__ import annotations

import re

from extractor_engine import normalize_product_name_for_quote


EL_COLOR_SUFFIX = "\n\n1.定位问题：——\n2.印花朝向问题：——\n3.条纹/格子款式是否对条/对格：——"


def _clean(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _ascii_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", _clean(value).lower())


def _contains_any(value: str, keywords: tuple[str, ...]) -> bool:
    value_l = value.lower()
    return any(keyword in value_l for keyword in keywords)


def _translate_composition(value: str) -> str:
    text = _clean(value)
    replacements = (
        (r"\bpolyamide\b", "尼龙"),
        (r"\bnylon\b", "尼龙"),
        (r"\belastane\b", "氨纶"),
        (r"\bspandex\b", "氨纶"),
        (r"\bpolyester\b", "涤纶"),
        (r"\bcotton\b", "棉"),
        (r"\bknitted\b", "针织"),
        (r"\bknit\b", "针织"),
        (r"\bgsm\b", "克"),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text, flags=re.I)
    text = re.sub(r"(\d+(?:[.,]\d+)?%)\s+", r"\1", text)
    text = re.sub(r"\s*,\s*", "，", text)
    text = re.sub(r"(\d+(?:[.,]\d+)?)\s*(?:克|g)\b", r"\1克", text, flags=re.I)
    return text.strip(" ，")


def normalize_el_product_name(text: str, raw_product_name: str = "", order_id: str = "") -> str:
    source = f"{raw_product_name}\n{text or ''}"
    source_l = source.lower()
    compact = _ascii_key(source)
    generic = normalize_product_name_for_quote(text, raw_product_name)
    has_upper_measure = any(keyword in source for keyword in ("胸杯", "胸围", "下胸围", "肩带", "领口", "夹围")) or any(
        keyword in compact for keyword in ("chest", "underchest", "neckopening", "armhole", "shoulderstrap", "frontdrop")
    )
    has_bottom_measure = any(keyword in source for keyword in ("裆", "前浪", "臀围", "大腿围", "腰围")) or any(
        keyword in compact for keyword in ("crotch", "rise", "hip", "thigh", "waist")
    )

    if "比基尼" in generic or "bikini" in source_l:
        garment = "比基尼套装"
    elif any(keyword in source for keyword in ("上衣", "分体", "两件套")) or _contains_any(source_l, ("two piece", "2-piece", "top and bottom")):
        garment = "分体泳装"
    elif "三角裤" in generic and not has_upper_measure:
        return generic
    elif has_upper_measure and has_bottom_measure:
        garment = "连体泳装"
    elif "连体" in generic:
        garment = "连体泳装"
    else:
        garment = "连体泳装" if "泳装" in generic or "swimsuit" in source_l else ""

    if any(keyword in source for keyword in ("大码",)) or _contains_any(source_l, ("plus size", "curve")):
        wearer = "女士大码"
    else:
        wearer = "女士"

    if any(keyword in source for keyword in ("固定杯", "模杯")) or _contains_any(source_l, ("fixed cup", "moulded cup", "molded cup", "schalen")):
        cup = "固定杯"
    elif any(keyword in source for keyword in ("塞杯", "可抽杯垫", "可拆杯垫")) or _contains_any(source_l, ("removable pad", "insert pad")):
        cup = "塞杯"
    elif "固定杯" in generic:
        cup = "固定杯"
    elif "塞杯" in generic:
        cup = "塞杯"
    elif garment:
        cup = "固定杯"
    else:
        cup = ""

    if garment and cup:
        return f"{wearer}{cup}{garment}"
    if garment:
        return f"{wearer}{garment}"
    return generic


def format_el_fabric_quality(text: str, raw_fabric_quality: str = "") -> str:
    source = f"{raw_fabric_quality}\n{text or ''}"
    source_l = source.lower()
    lines = []

    main_matches = re.findall(
        r"(\d+(?:[.,]\d+)?%\s*(?:polyamide|nylon|polyester|cotton)[^;\n]*?\d+(?:[.,]\d+)?%\s*(?:elastane|spandex|polyamide|nylon|polyester|cotton)(?:[^;\n]*?(?:\d+(?:[.,]\d+)?\s*(?:gsm|g|克)))?)",
        source,
        flags=re.I,
    )
    main = next((match for match in main_matches if "100%" not in match.lower() or "polyester" not in match.lower()), "")
    if main:
        main_text = _translate_composition(main)
        if not re.search(r"针织|泡泡布|皱皱布|提花", main_text):
            main_text = f"针织，{main_text}"
        lines.append(f"大身前后片：{main_text}")
    elif _contains_any(source_l, ("last season style", "same as last season")):
        style_match = re.search(r"(?:style|款)\s*[:：]?\s*([A-Z0-9-]+)", source, flags=re.I)
        style_note = f" {style_match.group(1)}" if style_match else ""
        lines.append(f"大身前后片：同上一季款{style_note}，成分克重请人工确认")

    if re.search(r"100%\s*(?:polyester|涤纶)", source, flags=re.I):
        lines.append("胸杯牛奶丝+前后片内衬：针织，100%涤纶，100克")

    return "\n\n".join(lines) if lines else raw_fabric_quality


def format_el_color_print(text: str, raw_color_print: str = "") -> str:
    source = f"{raw_color_print}\n{text or ''}"
    source_l = source.lower()

    if _contains_any(source_l, ("stripe", "striped")) or any(keyword in source for keyword in ("条纹",)):
        body = "大身前后片：泡泡布条纹"
    elif _contains_any(source_l, ("check", "checked", "plaid")) or any(keyword in source for keyword in ("格子", "色织格子")):
        body = "大身前后片：色织格子"
    elif _contains_any(source_l, ("all over print", "aop", "digital print", "print")) or any(keyword in source for keyword in ("满印", "数码印花", "印花", "图案")):
        body = "大身前后片：满印数码印花"
    elif raw_color_print:
        body = raw_color_print.splitlines()[0].strip()
    else:
        body = "大身前后片：净色"

    lining = "胸杯牛奶丝+前后片内衬：净色，配色"
    return f"{body}\n{lining}{EL_COLOR_SUFFIX}"


def apply_el_quote_rules(text: str, data: dict) -> dict:
    result = dict(data)
    result["product_name"] = normalize_el_product_name(
        text,
        result.get("product_name", ""),
        result.get("order_id", ""),
    )
    fabric_quality = format_el_fabric_quality(text, result.get("fabric_quality", ""))
    if fabric_quality:
        result["fabric_quality"] = fabric_quality
    result["color_print"] = format_el_color_print(text, result.get("color_print", ""))
    return result
