import logging
import os
import platform
from pathlib import Path

from typing import Tuple

from py4lo_commons import init_logger
from py4lo_helper import unohelper
import py4lo_helper
from ah4lo import AH4LO
from lo_helper import FakeProvider
from py4lo_dialogs import message_box, MessageBoxType

try:
    # noinspection PyUnresolvedReferences
    from com.sun.star.task import XJobExecutor
except (ModuleNotFoundError, ImportError):
    class XJobExecutor:
        pass


IMPLEMENTATION_NAME = "com.github.jferard.AccessibilityHelper4LO"
system = platform.system()
if system == "Windows":
    try:
        LOG_PATH = Path(os.environ["appdata"]) / "AccessibilityHelper4LO.log"
    except KeyError:
        LOG_PATH = None
elif system == "Linux":
    LOG_PATH = Path("/var/log/AccessibilityHelper4LO.log")
else:
    LOG_PATH = None


class AccessibilityHelper4LO(unohelper.Base, XJobExecutor):
    _inited = False
    _logger = logging.getLogger(__name__)

    def __init__(self, component_ctx):
        self._component_ctx = component_ctx
        self._oDesktop = self._component_ctx.getByName(
            "/singletons/com.sun.star.frame.theDesktop")
        self._oDoc = self._oDesktop.getCurrentComponent()
        py4lo_helper.provider = FakeProvider(component_ctx)

        if not AccessibilityHelper4LO._inited:
            init_logger(logging.getLogger(), LOG_PATH)
            self._logger.debug("Start of %s", self.__class__.__name__)
            AccessibilityHelper4LO._inited = True

        self._service_manager = self._component_ctx.getServiceManager()
        self._logger.debug("New %s instance", self.__class__.__name__)

    # XJobExecutor / void 	trigger ([in] string Event)
    def trigger(self, event: str):
        self._logger.debug("Function call: %s", event)
        # noinspection PyBroadException
        try:
            self._trigger(event.strip())
        except Exception:
            self._logger.exception("General error")

    # XServiceName /
    def getServiceName(self) -> str:
        return IMPLEMENTATION_NAME

    # XServiceInfo / string 	getImplementationName ()
    def getImplementationName(self) -> str:
        return IMPLEMENTATION_NAME

    # XServiceInfo / boolean 	supportsService ([in] string ServiceName)
    def supportsService(self, service_name: str) -> bool:
        return service_name in self.getSupportedServiceNames()

    # XServiceInfo / sequence< string > 	getSupportedServiceNames ()
    def getSupportedServiceNames(self) -> Tuple[str, ...]:
        return IMPLEMENTATION_NAME,

    def _trigger(self, func_name: str):
        try:
            func = METHOD_BY_NAME[func_name]
            func(self)
        except KeyError:
            message_box(
                "Missing function",
                "Function `{}` is missing".format(func_name),
                MessageBoxType.ERRORBOX)

    def run_calc(self):
        AH4LO(self._component_ctx, self._oDoc).run_calc()

    def run_writer(self):
        AH4LO(self._component_ctx, self._oDoc).run_writer()

METHOD_BY_NAME = {
    "run_calc": AccessibilityHelper4LO.run_calc,
    "run_writer": AccessibilityHelper4LO.run_writer
}

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    AccessibilityHelper4LO, IMPLEMENTATION_NAME, ('com.sun.star.task.Job',))
