import logging

import lo_helper
from ah4lo_dialogs import AH4LODialogs


class AH4LO:
    _logger = logging.getLogger(__name__)

    def __init__(self, component_ctx, oDoc):
        self._component_ctx = component_ctx
        self._oDoc = oDoc

    def run_calc(self):
        self._logger.debug("oDoc %s", self._oDoc.Title)
        lo_dialogs = AH4LODialogs(lo_helper.get_lang())
        oDialogControl = lo_dialogs.create_calc_control(self._oDoc)
        oDialogControl.setVisible(True)  # execute()

    def run_writer(self):
        self._logger.debug("oDoc %s", self._oDoc.Title)
        lo_dialogs = AH4LODialogs(lo_helper.get_lang())
        oDialogControl = lo_dialogs.create_writer_control(self._oDoc)
        oDialogControl.setVisible(True)  # execute()
