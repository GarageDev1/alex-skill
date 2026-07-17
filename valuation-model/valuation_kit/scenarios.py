"""三种情景的参数选择和工作表接口。"""

CASES = ("Bear", "Base", "Bull")


def nested_if_switch(sw_cell, case_names=CASES):
    formula = str(len(case_names))
    for index in range(len(case_names) - 2, -1, -1):
        formula = f'IF({sw_cell}="{case_names[index]}",{index + 1},{formula})'
    return "=" + formula


def select_case(values, case):
    if case not in CASES:
        raise ValueError(f"未知情景：{case}")
    return values[case]


def write_multiple_assumptions(*args, **kwargs):
    from .sheets import write_multiple_assumptions as writer
    return writer(*args, **kwargs)


def write_scenario_switch(*args, **kwargs):
    from .sheets import write_scenario_switch as writer
    return writer(*args, **kwargs)


def write_comparison(*args, **kwargs):
    from .sheets import write_comparison as writer
    return writer(*args, **kwargs)
