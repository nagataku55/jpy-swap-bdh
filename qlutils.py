"""カレンダー・日数計算・営業日調整・テナー解析などのQuantLib変換ヘルパー。

excel-fastapiプロジェクトのqlutils.pyを移植したもの(内容の変更なし)。
文字列コード(例: "TKB", "MF", "ACT365F")とQuantLibオブジェクトの対応表を
ここに集約することで、curves.py側は文字列ベースの設定(Excel/APIから渡しやすい形)
だけを扱えばよくなる。
"""
import QuantLib as ql

_CALENDAR_FACTORIES = {
    "TKB": ql.Japan,  # Tokyo Business day
    "TGT": ql.TARGET,  # TARGET (Euro)
    "NYB": lambda: ql.UnitedStates(ql.UnitedStates.GovernmentBond),  # New York Business day
    "LNB": ql.UnitedKingdom,  # London Business day
    "WKD": ql.WeekendsOnly,  # 土日のみ休日
    "NUL": ql.NullCalendar,  # 祝日なし(毎日営業日)
}


def get_calendar(name: str) -> ql.Calendar:
    """カンマ区切りで複数指定するとJointCalendar(JoinHolidays)を返す。

    例: "TKB" / "TGT,NYB"
    """
    names = [n.strip().upper() for n in name.split(",") if n.strip()]
    if not names:
        raise ValueError("calendar name is empty")

    calendars = []
    for n in names:
        factory = _CALENDAR_FACTORIES.get(n)
        if factory is None:
            raise ValueError(
                f"Unknown calendar: {n!r} (supported: {', '.join(sorted(_CALENDAR_FACTORIES))})"
            )
        calendars.append(factory())

    joint = calendars[0]
    for c in calendars[1:]:
        joint = ql.JointCalendar(joint, c, ql.JoinHolidays)
    return joint


_BDC_FACTORIES = {
    "MF": ql.ModifiedFollowing,
    "F": ql.Following,
    "P": ql.Preceding,
    "MP": ql.ModifiedPreceding,
    "U": ql.Unadjusted,
}


def get_business_day_convention(code: str) -> int:
    key = code.strip().upper()
    conv = _BDC_FACTORIES.get(key)
    if conv is None:
        raise ValueError(
            f"Unknown business day convention: {code!r} (supported: {', '.join(sorted(_BDC_FACTORIES))})"
        )
    return conv


def get_end_of_month(code: str) -> bool:
    key = code.strip().upper()
    if key == "STD":
        return False
    if key == "EOM":
        return True
    raise ValueError(f"Unknown eom option: {code!r} (supported: STD, EOM)")


_TENOR_ALIASES = {
    "ON": "1D",  # オーバーナイト
    "O/N": "1D",
}


def parse_period(tenor: str) -> ql.Period:
    key = tenor.strip().upper()
    key = _TENOR_ALIASES.get(key, key)
    try:
        return ql.Period(key)
    except RuntimeError as exc:
        raise ValueError(f"Invalid tenor: {tenor!r}") from exc


_RULE_FACTORIES = {
    "BACKWARD": ql.DateGeneration.Backward,
    "FORWARD": ql.DateGeneration.Forward,
    "ZERO": ql.DateGeneration.Zero,
}


def get_date_generation_rule(code: str) -> int:
    key = code.strip().upper()
    rule = _RULE_FACTORIES.get(key)
    if rule is None:
        raise ValueError(
            f"Unknown date generation rule: {code!r} (supported: {', '.join(sorted(_RULE_FACTORIES))})"
        )
    return rule


_DAY_COUNTER_FACTORIES = {
    "ACT360": ql.Actual360,
    "ACT365F": ql.Actual365Fixed,
    "ACTACT": lambda: ql.ActualActual(ql.ActualActual.ISDA),
    "30360": lambda: ql.Thirty360(ql.Thirty360.BondBasis),
}


def get_day_counter(code: str) -> ql.DayCounter:
    key = code.strip().upper()
    factory = _DAY_COUNTER_FACTORIES.get(key)
    if factory is None:
        raise ValueError(
            f"Unknown day count: {code!r} (supported: {', '.join(sorted(_DAY_COUNTER_FACTORIES))})"
        )
    return factory()


_FREQUENCY_FACTORIES = {
    "1Y": ql.Annual,
    "6M": ql.Semiannual,
    "3M": ql.Quarterly,
    "1M": ql.Monthly,
}


def get_frequency(code: str) -> int:
    key = code.strip().upper()
    freq = _FREQUENCY_FACTORIES.get(key)
    if freq is None:
        raise ValueError(
            f"Unknown frequency: {code!r} (supported: {', '.join(sorted(_FREQUENCY_FACTORIES))})"
        )
    return freq


_COMPOUNDING_FACTORIES = {
    "SIMPLE": ql.Simple,
    "COMPOUNDED": ql.Compounded,
    "CONTINUOUS": ql.Continuous,
}


def get_compounding(code: str) -> int:
    key = code.strip().upper()
    comp = _COMPOUNDING_FACTORIES.get(key)
    if comp is None:
        raise ValueError(
            f"Unknown compounding: {code!r} (supported: {', '.join(sorted(_COMPOUNDING_FACTORIES))})"
        )
    return comp


def approx_months(period: ql.Period) -> float:
    """1Y未満/以上の判定用(OISの支払頻度: 満期一括 vs 年次)の近似月数。"""
    unit = period.units()
    length = period.length()
    if unit == ql.Days:
        return length / 30.0
    if unit == ql.Weeks:
        return length * 7 / 30.0
    if unit == ql.Months:
        return float(length)
    if unit == ql.Years:
        return length * 12.0
    raise ValueError(f"Unsupported period unit in tenor: {period}")


_OVERNIGHT_INDEX_FACTORIES = {
    "TONAR": ql.Tonar,  # JPY TONA (Tokyo Overnight Average rate)
}


def get_overnight_index(code: str, handle: "ql.YieldTermStructureHandle | None" = None) -> ql.OvernightIndex:
    key = code.strip().upper()
    factory = _OVERNIGHT_INDEX_FACTORIES.get(key)
    if factory is None:
        raise ValueError(
            f"Unknown overnight index: {code!r} (supported: {', '.join(sorted(_OVERNIGHT_INDEX_FACTORIES))})"
        )
    if handle is None:
        return factory()
    return factory(handle)


_SWAP_POSITION_FACTORIES = {
    "PAY": ql.OvernightIndexedSwap.Payer,
    "RECEIVE": ql.OvernightIndexedSwap.Receiver,
}


def get_swap_position(code: str) -> int:
    key = code.strip().upper()
    position = _SWAP_POSITION_FACTORIES.get(key)
    if position is None:
        raise ValueError(
            f"Unknown swap position: {code!r} (supported: {', '.join(sorted(_SWAP_POSITION_FACTORIES))})"
        )
    return position
