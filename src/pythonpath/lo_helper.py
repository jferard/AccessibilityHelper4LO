import collections

try:
    # noinspection PyUnresolvedReferences
    from com.sun.star.util import NumberFormat
except (ModuleNotFoundError, ImportError):
    from mock_constants import NumberFormat

from py4lo_helper import create_uno_service, make_pv
from py4lo_typing import UnoRange


def get_lang() -> str:
    """
    :return: the language code
    """
    oConfigProvider = create_uno_service(
        "com.sun.star.configuration.ConfigurationProvider")
    oSetupL10N = oConfigProvider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationAccess",
        [make_pv("nodepath", "/org.openoffice.Setup/L10N")])
    return oSetupL10N.getByName("ooLocale")


def guess_format_id(oRange: UnoRange) -> int:
    if oRange is None:
        return 0
    if oRange.NumberFormat != 0:
        return oRange.NumberFormat

    oRangeAddress = oRange.RangeAddress
    row_count = min(1000, oRangeAddress.EndRow - oRangeAddress.StartRow + 1)

    oLimitedRange = oRange.getCellRangeByPosition(0, 0, 0, row_count - 1)
    data_array = oLimitedRange.DataArray

    non_empty_indices = [
        r
        for r, row in enumerate(data_array)
        if isinstance(row[0], float) or row[0]
    ]

    # essaie le format texte sur 1000 lignes ou moins
    float_count = 0
    text_count = 0
    for r in non_empty_indices:
        v = data_array[r][0]
        if isinstance(v, str):
            text_count += 1
        else:
            float_count += 1
    if text_count > float_count:  # plus de texte que de numérique
        return -1

    # plus long : format réel sur 1000 lignes ou moins
    counter = collections.Counter()
    for r in non_empty_indices[:100]:
        oCell = oRange.getCellByPosition(0, r)
        counter[oCell.NumberFormat] += 1

    return max(counter, key=lambda k: (counter[k], k), default=0)


type_by_format = {}


def get_type_id(oFormats, format_id: int) -> NumberFormat:
    try:
        return type_by_format[format_id]
    except KeyError:
        if format_id == -1:
            data_type = NumberFormat.TEXT
        else:
            oFormat = oFormats.getByKey(format_id)
            data_type = oFormat.Type & 0b111111111110
        type_by_format[format_id] = data_type
        return data_type


class FakeProvider:
    def __init__(self, component_ctx):
        self.service_manager = component_ctx.getServiceManager()
