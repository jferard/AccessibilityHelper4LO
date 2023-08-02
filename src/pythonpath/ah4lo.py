import lo_helper
from ah4lo_dialogs import AH4LODialogs
from py4lo_dialogs import MessageBoxType, message_box


class AH4LO:
    def __init__(self, component_ctx, oDoc):
        self._component_ctx = component_ctx
        self._oDoc = oDoc

    def run_calc(self):
        lo_dialogs = AH4LODialogs(lo_helper.get_lang())
        oDialogControl = lo_dialogs.create_calc_control(self._oDoc)
        oDialogControl.setVisible(True)  # execute()

    def run_writer(self):
        message_box(
            "AH4LO (Writer)", "TODO", MessageBoxType.MESSAGEBOX)
