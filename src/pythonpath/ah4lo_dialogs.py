import logging
from typing import Iterator

from ah4lo_data import DocumentNodeFactory
from ah4lo_lang import AH4LOLang
from ah4lo_tree import Node, Tree
from py4lo_dialogs import place_widget, Control, ControlModel
from py4lo_helper import create_uno_service, unohelper
from py4lo_typing import UnoControl, UnoControlModel, UnoSpreadsheet

try:
    # noinspection PyUnresolvedReferences
    from com.sun.star.awt import XKeyListener
except ImportError:
    class XKeyListener:
        pass

# Constants
DOWN_KEY = 0x400
UP_KEY = 0x401
RIGHT_KEY = 0x402
LEFT_KEY = 0x403
ENTER_KEY = 0x500


class ItemKeyListener(unohelper.Base, XKeyListener):
    _logger = logging.getLogger(__name__)

    def __init__(self, helper: "ScrollTreeHelper", oDialogControl: UnoControl):
        self.helper = helper
        self.oDialogControl = oDialogControl

    def keyPressed(self, _e):
        pass

    def keyReleased(self, e):
        self._logger.debug("Key %s", e.KeyCode)
        # noinspection PyBroadException
        try:
            state = self.helper.tree
            if e.KeyCode == DOWN_KEY:
                state.down()
            elif e.KeyCode == UP_KEY:
                state.up()
            elif e.KeyCode == RIGHT_KEY:
                state.right()
            elif e.KeyCode == LEFT_KEY:
                state.left()
            elif e.KeyCode == ENTER_KEY:
                state.enter()

            self.helper.place_lines(self.oDialogControl)
        except Exception:
            self._logger.exception("Key")


class ScrollTreeHelper:
    def __init__(self, root: Node, line_count: int, width: int, height: int,
                 prefix: str = "scroll_tree"):
        self.tree = Tree(root)
        self.line_count = line_count
        self.width = width
        self.height = height
        self.prefix = prefix

    def create_models(self, oDialogModel: UnoControlModel):
        for i in range(self.line_count):
            identifier = "{}{}".format(self.prefix, i)
            oTextModel = oDialogModel.createInstance(
                "com.sun.star.awt.UnoControlFixedTextModel")
            place_widget(oTextModel, 0, self.height * i, self.width,
                         self.height)
            oTextModel.Name = identifier
            oTextModel.Tabstop = True
            oTextModel.FontName = "Liberation Mono"
            if i == self.line_count // 2:
                oTextModel.FontWeight = 150
            oDialogModel.insertByName(identifier, oTextModel)

    def _get_controls(self, oDialogControl: UnoControl
                      ) -> Iterator[UnoControl]:
        for oControl in oDialogControl.getControls():
            if oControl.Model.Name.startswith(self.prefix):
                yield oControl

    def add_keys(self, oDialogControl: UnoControl):
        listener = ItemKeyListener(self, oDialogControl)
        for oControl in self._get_controls(oDialogControl):
            oControl.addKeyListener(listener)

    def place_lines(self, oDialogControl: UnoControl):
        controls = list(self._get_controls(oDialogControl))

        base_index = len(controls) // 2

        oTextControl = controls[base_index]
        oTextControl.Visible = True
        oTextControl.Model.Label = self.tree.text(self.tree.focus)
        oTextControl.setFocus()

        f = self.tree.focus
        for i in range(base_index - 1, -1, -1):
            oTextControl = controls[i]
            f = f.previous() if f else None
            if f is None:
                oTextControl.Visible = False
            else:
                oTextControl.Visible = True
                oTextControl.Model.Label = self.tree.text(f)

        f = self.tree.focus
        for i in range(base_index + 1, self.line_count):
            oTextControl = controls[i]
            f = f.next() if f else None
            if f is None:
                oTextControl.Visible = False
            else:
                oTextControl.Visible = True
                oTextControl.Model.Label = self.tree.text(f)


class AH4LODialogs:
    _logger = logging.getLogger(__name__)

    def __init__(self, lo_lang: str):
        self._ah4lo_lang = AH4LOLang.from_lang(lo_lang)

    def create_calc_control(self, oDoc: UnoSpreadsheet):
        doc_title = oDoc.Title
        sheet_count = oDoc.Sheets.Count
        text = self._ah4lo_lang.calc_window_title(
            doc_title, sheet_count)

        oDialogModel = create_uno_service(ControlModel.Dialog)
        oDialogModel.Title = text
        place_widget(oDialogModel, 100, 50, 500, 300)

        root = DocumentNodeFactory(self._ah4lo_lang, oDoc).get_root()
        helper = ScrollTreeHelper(root, 15, 400, 15)
        helper.create_models(oDialogModel)

        oDialogControl = create_uno_service(Control.Dialog)
        oDialogControl.setModel(oDialogModel)

        helper.add_keys(oDialogControl)
        helper.place_lines(oDialogControl)

        oDialogControl.setVisible(True)
        toolkit = create_uno_service(
            "com.sun.star.awt.Toolkit")
        oDialogControl.createPeer(toolkit, None)
        return oDialogControl
