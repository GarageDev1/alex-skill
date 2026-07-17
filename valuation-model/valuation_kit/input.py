"""统一输入的读取与校验。"""

from __future__ import annotations

import json
from pathlib import Path

from jsonschema import Draft202012Validator, FormatChecker


SCHEMA_PATH = Path(__file__).with_name("schema") / "input.schema.json"
METHODS = {"P/E", "P/B", "EV/Sales", "DCF", "SOTP"}
CASES = ("Bear", "Base", "Bull")


class InputValidationError(ValueError):
    pass


def _need(mapping, key, where):
    if key not in mapping:
        raise InputValidationError(f"{where} 缺少字段 {key}")
    return mapping[key]


def _series(series, size, where):
    if not isinstance(series, dict):
        raise InputValidationError(f"{where} 必须包含 values 和 source_id")
    values = _need(series, "values", where)
    _need(series, "source_id", where)
    if not isinstance(values, list) or len(values) != size:
        raise InputValidationError(f"{where}.values 必须有 {size} 项")
    if any(value is not None and not isinstance(value, (int, float)) for value in values):
        raise InputValidationError(f"{where}.values 只能填写数字或 null")


def validate_input(data):
    """校验结构、年份、币种和来源引用，失败时抛出 InputValidationError。"""
    if not isinstance(data, dict):
        raise InputValidationError("输入必须是 JSON 对象")
    schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
    errors = sorted(Draft202012Validator(schema, format_checker=FormatChecker()).iter_errors(data),
                    key=lambda error: list(error.absolute_path))
    if errors:
        error = errors[0]
        path = ".".join(str(part) for part in error.absolute_path) or "根对象"
        raise InputValidationError(f"{path}: {error.message}")
    for key in ("schema_version", "company", "currency", "years", "history", "drivers",
                "scenarios", "valuation_methods", "sources"):
        _need(data, key, "根对象")

    company = data["company"]
    currency = data["currency"]
    years = data["years"]
    sources = data["sources"]
    if currency.get("financial") != "USD":
        raise InputValidationError("currency.financial 必须为 USD")
    price_currency = _need(currency, "price", "currency")
    fx = _need(currency, "local_per_usd", "currency")
    if fx.get("currency") != price_currency or fx.get("value", 0) <= 0:
        raise InputValidationError("汇率币种必须与股价币种一致，且汇率必须大于 0")

    historical = _need(years, "historical", "years")
    forecast = _need(years, "forecast", "years")
    if not historical or not forecast:
        raise InputValidationError("历史年份和预测年份都不能为空")
    if historical != sorted(historical) or forecast != sorted(forecast):
        raise InputValidationError("年份必须升序排列")
    if historical[-1] >= forecast[0] or len(set(historical + forecast)) != len(historical + forecast):
        raise InputValidationError("历史年份与预测年份不能重叠或错位")

    for name in ("revenue", "net_income", "equity", "shares_millions", "price", "local_per_usd"):
        _series(_need(data["history"], name, "history"), len(historical), f"history.{name}")
    shares = _need(company, "shares_millions", "company")
    if shares.get("value", 0) <= 0:
        raise InputValidationError("company.shares_millions.value 必须大于 0")
    price = _need(company, "price", "company")
    if price.get("currency") != price_currency:
        raise InputValidationError("公司股价币种与 currency.price 不一致")
    net_debt = _need(company, "net_debt_usd_billions", "company")
    if not isinstance(net_debt.get("value"), (int, float)):
        raise InputValidationError("company.net_debt_usd_billions.value 必须是数字")

    if not isinstance(data["drivers"], list) or not data["drivers"]:
        raise InputValidationError("drivers 至少需要一项")
    for index, driver in enumerate(data["drivers"]):
        _series(_need(driver, "historical", f"drivers[{index}]"), len(historical),
                f"drivers[{index}].historical")
        _series(_need(driver, "forecast_base", f"drivers[{index}]"), len(forecast),
                f"drivers[{index}].forecast_base")

    for case in CASES:
        case_data = _need(data["scenarios"], case, "scenarios")
        adjustments = _need(case_data, "driver_adjustments", f"scenarios.{case}")
        margins = _need(case_data, "net_margin", f"scenarios.{case}")
        if len(adjustments) != len(data["drivers"]):
            raise InputValidationError(f"scenarios.{case}.driver_adjustments 数量与 drivers 不一致")
        if len(margins) != len(forecast):
            raise InputValidationError(f"scenarios.{case}.net_margin 数量与预测年份不一致")
        _need(case_data, "source_id", f"scenarios.{case}")

    methods = data["valuation_methods"]
    if not isinstance(methods, list) or len(methods) < 2:
        raise InputValidationError("valuation_methods 至少需要两种估值方法")
    names = set()
    for index, method in enumerate(methods):
        name = _need(method, "name", f"valuation_methods[{index}]")
        if name not in METHODS:
            raise InputValidationError(f"不支持的估值方法：{name}")
        if name in names:
            raise InputValidationError(f"估值方法重复：{name}")
        names.add(name)
        if name in {"P/E", "P/B", "EV/Sales"}:
            multiples = _need(method, "multiples", f"valuation_methods[{index}]")
            if set(multiples) != set(CASES) or any(not isinstance(v, (int, float)) for v in multiples.values()):
                raise InputValidationError(f"valuation_methods[{index}].multiples 必须包含 Bear/Base/Bull 数字")
        elif name == "DCF":
            assumptions = _need(method, "assumptions", f"valuation_methods[{index}]")
            for case in CASES:
                values = _need(assumptions, case, f"valuation_methods[{index}].assumptions")
                if len(values.get("fcf_margins", [])) != len(forecast):
                    raise InputValidationError(f"DCF {case} 的 fcf_margins 数量与预测年份不一致")
                if values.get("discount_rate", 0) <= values.get("terminal_growth", 0):
                    raise InputValidationError(f"DCF {case} 的折现率必须高于永续增长率")
        else:
            assumptions = _need(method, "assumptions", f"valuation_methods[{index}]")
            for case in CASES:
                parts = _need(assumptions, case, f"valuation_methods[{index}].assumptions")
                if not isinstance(parts, list) or not parts:
                    raise InputValidationError(f"SOTP {case} 至少需要一个分部")
                for part in parts:
                    if part.get("metric") not in {"revenue", "net_income", "equity"}:
                        raise InputValidationError(f"SOTP {case} 的 metric 不受支持")
                    if part.get("target_year") not in forecast or not isinstance(part.get("multiple"), (int, float)):
                        raise InputValidationError(f"SOTP {case} 的年份或倍数无效")
        if method.get("target_year") not in forecast:
            raise InputValidationError(f"valuation_methods[{index}].target_year 必须属于预测年份")
        _need(method, "source_id", f"valuation_methods[{index}]")

    source_ids = set()
    source_ids.update(item["source_id"] for item in data["history"].values())
    source_ids.update((company["price"]["source_id"], shares["source_id"],
                       net_debt["source_id"], fx["source_id"]))
    for driver in data["drivers"]:
        source_ids.update((driver["historical"]["source_id"], driver["forecast_base"]["source_id"]))
    source_ids.update(data["scenarios"][case]["source_id"] for case in CASES)
    source_ids.update(method["source_id"] for method in methods)
    missing = sorted(source_ids - set(sources))
    if missing:
        raise InputValidationError(f"sources 缺少引用项：{missing}")
    for source_id, source in sources.items():
        for key in ("title", "publisher", "url", "as_of", "currency_basis"):
            if not source.get(key):
                raise InputValidationError(f"sources.{source_id} 缺少字段 {key}")
        if not str(source["url"]).startswith(("https://", "http://")):
            raise InputValidationError(f"sources.{source_id}.url 必须是完整网址")
    return data


def load_input(path):
    input_path = Path(path)
    try:
        data = json.loads(input_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise InputValidationError(f"无法读取输入文件 {input_path}: {exc}") from exc
    return validate_input(data)
