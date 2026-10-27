import os
import re
import json
from typing import Dict, List, Optional
from dotenv import load_dotenv
from rapidfuzz import fuzz
from groq import Groq

load_dotenv()

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()

FIELD_KNOWLEDGE_BASE: Dict[str, List[str]] = {
    "eup_inc_vat": [
        "end user price include vat",
        "end user price incl vat",
        "end user price including vat",
        "end user price in vat",
        "eup in vat",
        "eup inc vat",
        "eup include vat",
        "eup vat",
        "harga eup include vat",
    ],
    "eup_ex_vat": [
        "end user price ex vat",
        "end user price excl vat",
        "end user price exclude vat",
        "end user price excluding vat",
        "eup ex vat",
        "eup excl vat",
        "eup exclude vat",
        "eup without vat",
        "harga eup ex vat",
    ],
    "payment_rate": [
        "payment rate",
        "pay rate",
        "payment %",
        "rate payment",
        "payment fee rate",
    ],
    "payment_cost": [
        "payment cost",
        "pay cost",
        "payment fee",
        "biaya payment",
        "cost of payment",
    ],
    "unit_price": [
        "unit price",
        "price per unit",
        "price unit",
        "harga satuan",
        "unit price idr",
    ],
    "cogs": [
        "cogs",
        "cost of goods sold",
        "cost of goods",
        "hpp",
        "harga pokok penjualan",
    ],
    "margin": [
        "margin",
        "profit margin",
        "gross margin",
        "margin %",
        "margin rate",
    ],
}

ACRONYM_MAP = {
    "eup": "end user price",
    "inc": "include",
    "in": "include",
    "ex": "exclude",
    "excl": "exclude",
    "hpp": "cost of goods sold",
}

def clean_text(text: str) -> str:
    cleaned = text.lower().strip()
    cleaned = re.sub(r"[^\w\s%]", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned

def expand_acronyms(text: str) -> str:
    words = clean_text(text).split()
    expanded_words = [ACRONYM_MAP.get(word, word) for word in words]
    return " ".join(expanded_words)

def match_single_header_fuzzy(raw_header: str, threshold: float = 60.0) -> Optional[str]:
    cleaned = clean_text(raw_header)
    expanded = expand_acronyms(raw_header)

    best_field = None
    best_score = 0.0

    for standard_field, aliases in FIELD_KNOWLEDGE_BASE.items():
        for alias in aliases:
            cleaned_alias = clean_text(alias)
            score_direct = fuzz.token_sort_ratio(cleaned, cleaned_alias)
            score_expanded = fuzz.token_sort_ratio(expanded, cleaned_alias)
            current_max = max(score_direct, score_expanded)

            if current_max > best_score:
                best_score = current_max
                best_field = standard_field

    if best_score >= threshold:
        return best_field
    return None

def map_headers_with_llm(raw_headers: List[str]) -> Optional[Dict[str, Optional[str]]]:
    api_key = os.getenv("GROQ_API_KEY", "").strip()
    if not api_key:
        return None

    client = Groq(api_key=api_key)

    prompt = f"""You are a data engineering AI assistant.
Your task is to map messy/abbreviated spreadsheet column headers into standard schema field names.

Standard fields and their meanings:
- eup_inc_vat: End User Price including VAT / Tax (examples: "end user price include vat", "EUP in VAT", "EUP inc VAT")
- eup_ex_vat: End User Price excluding VAT (examples: "end user price ex vat", "EUP ex VAT")
- payment_rate: Payment fee rate percentage (examples: "payment rate", "pay rate")
- payment_cost: Payment fee cost (examples: "payment cost", "biaya payment")
- unit_price: Unit price / base price (examples: "unit price", "price per unit")
- cogs: Cost of Goods Sold (examples: "cogs", "cost of goods sold", "hpp")
- margin: Profit margin (examples: "margin", "gross margin")

Headers to map:
{json.dumps(raw_headers)}

Return a strict JSON object with a single key "mapping" containing a dictionary where each key is the exact raw header string and the value is either one of the 7 standard field names or null if unrelated.
Example format:
{{
  "mapping": {{
    "EUP in VAT": "eup_inc_vat",
    "unknown": null
  }}
}}"""

    try:
        response = client.chat.completions.create(
            model="openai/gpt-oss-120b",
            messages=[
                {"role": "system", "content": "You are a JSON schema mapping assistant. Output valid JSON only."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.0
        )
        content = response.choices[0].message.content
        data = json.loads(content)
        mapping = data.get("mapping", {})
        
        valid_fields = set(FIELD_KNOWLEDGE_BASE.keys())
        result: Dict[str, Optional[str]] = {}
        for h in raw_headers:
            val = mapping.get(h)
            if val in valid_fields:
                result[h] = val
            else:
                result[h] = None
        return result
    except Exception as e:
        return None

def map_headers(raw_headers: List[str]) -> Dict[str, Optional[str]]:
    llm_result = map_headers_with_llm(raw_headers)
    if llm_result is not None:
        return llm_result

    result: Dict[str, Optional[str]] = {}
    assigned_fields = set()

    for header in raw_headers:
        matched = match_single_header_fuzzy(header)
        if matched and matched not in assigned_fields:
            result[header] = matched
            assigned_fields.add(matched)
        elif matched and matched in assigned_fields:
            result[header] = matched
        else:
            result[header] = None

    return result
