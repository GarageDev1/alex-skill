"""常用估值方法。所有金额均以十亿美元计。"""

from .currency import local_price_from_usd_per_share, usd_per_share
from .sheets import write_scenario_valuation


SUPPORTED_METHODS = ("P/E", "P/B", "EV/Sales", "DCF", "SOTP")


def pe_price(net_income, shares, multiple, fx):
    return local_price_from_usd_per_share(usd_per_share(net_income, shares) * multiple, fx)


def pb_price(equity, shares, multiple, fx):
    return local_price_from_usd_per_share(usd_per_share(equity, shares) * multiple, fx)


def ev_sales_price(revenue, net_debt, shares, multiple, fx):
    equity_value = revenue * multiple - net_debt
    return local_price_from_usd_per_share(usd_per_share(equity_value, shares), fx)


def dcf_value(cash_flows, discount_rate, terminal_growth):
    if discount_rate <= terminal_growth:
        raise ValueError("折现率必须高于永续增长率")
    if not cash_flows:
        raise ValueError("现金流不能为空")
    value = sum(flow / (1 + discount_rate) ** year for year, flow in enumerate(cash_flows, 1))
    terminal = cash_flows[-1] * (1 + terminal_growth) / (discount_rate - terminal_growth)
    return value + terminal / (1 + discount_rate) ** len(cash_flows)


def sotp_value(parts, net_debt=0):
    """parts 为包含 value 或 metric×multiple 的字典列表。"""
    total = 0.0
    for part in parts:
        total += part.get("value", part.get("metric", 0) * part.get("multiple", 0))
    return total - net_debt
