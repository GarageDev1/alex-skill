"""币种、金额单位和每股口径换算。"""


def local_billions_to_usd_billions(value, local_per_usd):
    """本币十亿元换算为十亿美元。"""
    if value is None:
        return None
    if local_per_usd <= 0:
        raise ValueError("汇率必须大于 0")
    return value / local_per_usd


def local_trillions_to_usd_billions(value, local_per_usd):
    """本币万亿元换算为十亿美元。"""
    if value is None:
        return None
    return local_billions_to_usd_billions(value * 1000, local_per_usd)


def usd_per_share(amount_usd_billions, shares_millions):
    if shares_millions <= 0:
        raise ValueError("股本必须大于 0")
    return amount_usd_billions * 1000 / shares_millions


def local_price_from_usd_per_share(value, local_per_usd):
    if local_per_usd <= 0:
        raise ValueError("汇率必须大于 0")
    return value * local_per_usd


def market_cap_usd_billions(price_local, shares_millions, local_per_usd):
    if local_per_usd <= 0 or shares_millions <= 0:
        raise ValueError("汇率和股本必须大于 0")
    return price_local * shares_millions / local_per_usd / 1000


def uy(value, fx):
    """兼容旧接口：本币万亿元换算为十亿美元。"""
    result = local_trillions_to_usd_billions(value, fx)
    return None if result is None else round(result, 2)
