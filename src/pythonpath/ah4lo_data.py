import logging
from typing import Optional, cast, Dict, List, NewType

from ah4lo_lang import AH4LOLang
from ah4lo_tree import Node, NodeBuilder, Action
from lo_helper import guess_format_id, get_type_id, extract_values
from py4lo_helper import get_used_range, to_iter
from py4lo_typing import (UnoSpreadsheet, UnoRange, UnoSheet, UnoService,
                          UnoController)


##############################
# CALC
##############################


class CalcDocumentNodeFactory:
    """
    A factory to build the document tree
    """

    def __init__(self, ah4lo_lang: AH4LOLang, oDoc: UnoSpreadsheet):
        self.ah4lo_lang = ah4lo_lang
        self.oDoc = oDoc
        self.oSheets = self.oDoc.Sheets
        self.oController = self.oDoc.CurrentController

    def get_root(self) -> Node:
        root_node = NodeBuilder(self.ah4lo_lang.sheets(self.oSheets.Count))
        for i in range(self.oSheets.Count):
            oSheet = self.oSheets.getByIndex(i)
            sheet_node = SheetNodeFactory(self.ah4lo_lang, self.oDoc,
                                          oSheet).get_root()
            root_node.append_child(sheet_node)
        root_node.freeze_as_root()
        return cast(Node, root_node)


class SheetNodeFactory:
    _logger = logging.getLogger(__name__)

    def __init__(self, ah4lo_lang: AH4LOLang, oDoc: UnoSpreadsheet,
                 oSheet: UnoSheet):
        self.ah4lo_lang = ah4lo_lang
        self.oDoc = oDoc
        self.oSheet = oSheet
        self.oSheets = self.oDoc.Sheets
        self.oController = self.oDoc.CurrentController

    def get_root(self) -> NodeBuilder:
        def action(oController=self.oController,
                   oSheet=self.oSheet):  # capture
            oController.ActiveSheet = oSheet

        name = self._get_sheet_description()
        sheet_node = NodeBuilder(name, action)
        oRange = get_used_range(self.oSheet)
        range_address = oRange.RangeAddress
        column_count = range_address.EndColumn - range_address.StartColumn + 1
        row_count = range_address.EndRow - range_address.StartRow + 1
        text = self.ah4lo_lang.used_range(column_count, row_count)
        range_node = NodeBuilder(text)
        sheet_node.append_child(range_node)
        columns_node = self.get_columns(oRange)
        sheet_node.append_child(columns_node)

        annotations_node = self.get_annotations()
        if annotations_node:
            sheet_node.append_child(annotations_node)

        dialogs_node = self.get_dialogs()
        if dialogs_node:
            sheet_node.append_child(dialogs_node)

        charts_node = self.get_charts()
        if charts_node:
            sheet_node.append_child(charts_node)

        data_pilot_tables_node = self.get_data_pilot_tables()
        if data_pilot_tables_node:
            sheet_node.append_child(data_pilot_tables_node)

        return sheet_node

    def _get_sheet_description(self):
        sheet_name = self.oSheet.Name
        is_hidden = not self.oSheet.IsVisible
        is_protected = self.oSheet.isProtected()
        return self.ah4lo_lang.sheet_description(
            sheet_name, is_hidden, is_protected)

    def get_columns(self, oRange: UnoRange) -> NodeBuilder:
        nodes = []
        oColumns = oRange.Columns
        for c in range(oColumns.Count):
            oColumn = oColumns.getByIndex(c)
            column_name = oColumn.getCellByPosition(0, 0).String
            if not column_name.strip():
                column_name = self.ah4lo_lang.empty_word

            oRangeWithoutHeader = oColumn.getCellRangeByPosition(
                0, 1, 0, oColumn.RangeAddress.EndRow)
            format_id = guess_format_id(oRangeWithoutHeader)
            type_id = get_type_id(self.oDoc.NumberFormats, format_id)
            type_name = self.ah4lo_lang.get_type_name(type_id)
            text = "{} {}".format(column_name, type_name)

            oController = self.oController

            def action(oController=oController, oColumn=oColumn):
                oController.select(oColumn)

            node = NodeBuilder(text, action)
            nodes.append(node)

        columns_node = NodeBuilder(self.ah4lo_lang.columns(len(nodes)))
        columns_node.extend_children(nodes)
        return columns_node

    def get_dialogs(self) -> Optional[NodeBuilder]:
        oForms = self.oSheet.DrawPage.Forms

        nodes = []
        for i in range(oForms.Count):
            oForm = oForms.getByIndex(i)
            for j in range(oForm.Count):
                oControlModel = oForm.getByIndex(j)
                name = self._extract_name(oControlModel.ServiceName)
                # CheckBox
                # ComboBox
                # CommandButton
                # CurrencyField
                # DatabaseCheckBox
                # DatabaseComboBox
                # DatabaseCurrencyField
                # DatabaseDateField
                # DatabaseFormattedField
                # DatabaseImageControl
                # DatabaseListBox
                # DatabaseNumericField
                # DatabasePatternField
                # DatabaseRadioButton
                # DatabaseTextField
                # DatabaseTimeField
                # DataForm
                # DateField
                # FileControl
                # FixedText
                # Form
                # FormattedField
                # GridControl
                # GroupBox
                # HiddenControl
                # HTMLForm
                # ImageButton
                # ListBox
                # NavigationToolBar
                # NumericField
                # PatternField
                # RadioButton
                # RichTextControl
                # ScrollBar
                # SpinButton
                # SubmitButton
                # TextField
                # TimeField
                text = "{}, {}".format(name, oControlModel.Label)

                def action(oController=self.oController,
                           oSheet=self.oSheet, oControlModel=oControlModel):
                    oController.ActiveSheet = oSheet
                    try:
                        oControl = self.oController.getControl(oControlModel)
                    except Exception:
                        self._logger.exception("control %s",
                                               oControlModel.Name)
                    else:
                        oControl.setFocus()

                node = NodeBuilder(text, action)
                nodes.append(node)

        if not nodes:
            return None

        dialogs_node = NodeBuilder(self.ah4lo_lang.dialogs(len(nodes)))
        dialogs_node.extend_children(nodes)

        return dialogs_node

    def _extract_name(self, service_name: str) -> str:
        parts = service_name.rsplit(".")
        return parts[-1]

    def get_annotations(self) -> Optional[NodeBuilder]:
        oAnnotations = self.oSheet.Annotations

        ranges_by_annotation = {}
        for i in range(oAnnotations.Count):
            oAnnotation = oAnnotations.getByIndex(i)
            pos = oAnnotation.Position
            text = oAnnotation.Text.String.strip()
            oCell = self.oSheet.getCellByPosition(pos.Column, pos.Row)
            if text in ranges_by_annotation:
                oRanges = ranges_by_annotation[text]
            else:
                oRanges = self.oDoc.createInstance(
                    "com.sun.star.sheet.SheetCellRanges")
                ranges_by_annotation[text] = oRanges

            oRanges.addRangeAddress(oCell.RangeAddress, True)

        if not ranges_by_annotation:
            return None

        annotations_node = NodeBuilder(
            self.ah4lo_lang.annotations(len(ranges_by_annotation)))
        for annotation, oRanges in ranges_by_annotation.items():
            node = NodeBuilder(
                "'{}' : {}".format(annotation, oRanges.RangeAddressesAsString))
            annotations_node.append_child(node)
        return annotations_node

    def get_charts(self) -> Optional[NodeBuilder]:
        oCharts = self.oSheet.Charts

        nodes = []
        for oChart in to_iter(oCharts):
            oTitle = oChart.EmbeddedObject.Title
            if oTitle:
                string = oTitle.String
            else:
                string = self.ah4lo_lang.anonymous_chart_word
            chart_node = NodeBuilder(string)
            nodes.append(chart_node)

        if not nodes:
            return None

        charts_node = NodeBuilder(self.ah4lo_lang.charts(len(nodes)))
        charts_node.extend_children(nodes)
        return charts_node

    def get_data_pilot_tables(self) -> Optional[NodeBuilder]:
        oDataPilotTables = self.oSheet.DataPilotTables
        nodes = []
        for oDataPilotTable in to_iter(oDataPilotTables):
            range_address = oDataPilotTable.SourceRange
            oSourceSheet = self.oDoc.Sheets.getByIndex(range_address.Sheet)
            oSourceRange = oSourceSheet.getCellRangeByPosition(
                range_address.StartColumn,
                range_address.StartRow,
                range_address.EndColumn,
                range_address.EndRow,
            )
            string = "Source {}".format(oSourceRange.AbsoluteName)
            data_pilot_table_node = NodeBuilder(string)
            nodes.append(data_pilot_table_node)

        if not nodes:
            return None

        data_pilot_tables_node = NodeBuilder(
            self.ah4lo_lang.dynamic_tables(len(nodes)))
        data_pilot_tables_node.extend_children(nodes)
        return data_pilot_tables_node


##############################
# WRITER
##############################
# services
TEXT_GRAPHIC_OBJECT_SERVICE_NAME = "com.sun.star.text.TextGraphicObject"

SHAPE_SERVICE_NAME = "com.sun.star.drawing.Shape"

TEXT_FRAME_SERVICE_NAME = "com.sun.star.text.TextFrame"

TEXT_EMBEDDED_OBJECT_SERVICE_NAME = "com.sun.star.text.TextEmbeddedObject"

BASE_FRAME_SERVICE_NAME = "com.sun.star.text.BaseFrame"

TEXT_TABLE_SERVICE_NAME = "com.sun.star.text.TextTable"

PARAGRAPH_SERVICE_NAME = "com.sun.star.text.Paragraph"

# Types
Paragraph = NewType("Paragraph", UnoService)
XShape = NewType("XShape", UnoService)
TextRange = NewType("TextRange", UnoService)


class DrawingNodeFactory:
    _logger = logging.getLogger(__name__)

    def __init__(self, ah4lo_lang: AH4LOLang, oDoc: UnoSpreadsheet,
                 drawings_by_paragraph: Dict[Paragraph, List[XShape]]):
        self.ah4lo_lang = ah4lo_lang
        self.oDoc = oDoc
        self.drawings_by_paragraph = drawings_by_paragraph

    def find_drawings(self, oParagraph: Paragraph) -> List[XShape]:
        return self.drawings_by_paragraph.get(oParagraph, [])

    def create_drawing_node(self, oDrawing: XShape, action: Optional[Action]
                            ) -> NodeBuilder:
        if oDrawing.supportsService(
                TEXT_EMBEDDED_OBJECT_SERVICE_NAME):
            self._logger.warning("TODO: Embedded object %s",
                                 repr(oDrawing.Component))
            value = self.ah4lo_lang.embedded_object(oDrawing.Name)
            return NodeBuilder(value, action)
        elif oDrawing.supportsService(TEXT_FRAME_SERVICE_NAME):
            value = self.ah4lo_lang.text_frame(oDrawing.Name)
            tf_node = NodeBuilder(
                value, action)
            WriterRangeContentBuilder(
                self.ah4lo_lang, self.oDoc, oDrawing, tf_node,
                self
            ).build()
            return tf_node
        elif oDrawing.supportsService(
                TEXT_GRAPHIC_OBJECT_SERVICE_NAME):
            value = self.ah4lo_lang.graphic_object(
                oDrawing.Name)
            return NodeBuilder(value, action)
        elif oDrawing.supportsService(SHAPE_SERVICE_NAME):
            value = self.ah4lo_lang.shape(oDrawing.Name)
            return NodeBuilder(value, action)
        else:
            self._logger.warning(
                "Unkown drawing: %s", repr(oDrawing))
            value = self.ah4lo_lang.unknown_drawing(oDrawing.Name)
            return NodeBuilder(value, action)


class WriterDocumentNodeFactory:
    _logger = logging.getLogger(__name__)

    def __init__(self, ah4lo_lang: AH4LOLang, oDoc: UnoSpreadsheet):
        self.ah4lo_lang = ah4lo_lang
        self.oDoc = oDoc
        self.oParagraphStyles = self.oDoc.StyleFamilies.ParagraphStyles
        self.oNumberingStyles = self.oDoc.StyleFamilies.NumberingStyles

    def get_root(self) -> Node:
        oProperties = self.oDoc.DocumentProperties
        drawings_by_paragraph = {}
        for oDrawPage in to_iter(self.oDoc.DrawPages):
            for oDrawing in to_iter(oDrawPage):
                oAnchor = oDrawing.Anchor
                if oAnchor is not None:
                    drawings_by_paragraph.setdefault(
                        oAnchor.TextParagraph, []).append(oDrawing)
        drawing_node_factory = DrawingNodeFactory(
            self.ah4lo_lang, self.oDoc, drawings_by_paragraph)

        root_node = NodeBuilder(oProperties.Title)

        information_node = self._create_informations_node()
        root_node.append_child(information_node)

        content_node = self._create_content_node(drawing_node_factory)
        root_node.append_child(content_node)

        orphans_node = self._create_orphans_node(drawing_node_factory)
        if orphans_node:
            root_node.append_child(orphans_node)

        root_node.freeze_as_root()
        return root_node

    def _create_informations_node(self) -> NodeBuilder:
        oProperties = self.oDoc.DocumentProperties
        information_node = NodeBuilder(self.ah4lo_lang.informations())
        if oProperties.Author.strip():
            value = self.ah4lo_lang.writer_author(oProperties.Author)
            author_node = NodeBuilder(value)
            information_node.append_child(author_node)
        if oProperties.Subject.strip():
            value = self.ah4lo_lang.writer_subject(oProperties.Subject)
            subject_node = NodeBuilder(value)
            information_node.append_child(subject_node)
        if oProperties.Description.strip():
            value = self.ah4lo_lang.writer_description(oProperties.Description)
            description_node = NodeBuilder(value)
            information_node.append_child(description_node)

        oStatistics = oProperties.DocumentStatistics
        page_count, paragraph_count, word_count = extract_values(
            oStatistics, ("PageCount", "ParagraphCount", "WordCount")
        )
        statistics_node = NodeBuilder(
            self.ah4lo_lang.statistics(
                page_count, paragraph_count, word_count))
        information_node.append_child(statistics_node)
        return information_node

    def _create_content_node(
            self, drawing_node_factory: DrawingNodeFactory
    ) -> NodeBuilder:
        oCursor = self.oDoc.Text.createTextCursor()
        oCursor.gotoStart(False)
        oCursor.gotoEnd(True)

        content_node = NodeBuilder(self.ah4lo_lang.content())
        return WriterRangeContentBuilder(
            self.ah4lo_lang, self.oDoc, oCursor, content_node,
            drawing_node_factory).build()

    def _create_orphans_node(self, drawing_node_factory: DrawingNodeFactory
                             ) -> Optional[NodeBuilder]:
        orphans = []
        for oDrawPage in to_iter(self.oDoc.DrawPages):
            for oDrawing in to_iter(oDrawPage):
                oAnchor = oDrawing.Anchor
                if oAnchor is None:
                    orphans.append(oDrawing)
        if orphans:
            orphans_node = NodeBuilder("Orphan drawings")
            for orphan in orphans:
                orphan_node = drawing_node_factory.create_drawing_node(orphan,
                                                                       None)
                orphans_node.append_child(orphan_node)
        else:
            orphans_node = None
        return orphans_node


class WriterRangeContentBuilder:
    _logger = logging.getLogger(__name__)

    def __init__(self, ah4lo_lang: AH4LOLang, oDoc: UnoSpreadsheet,
                 oTextRange: TextRange, content_node: NodeBuilder,
                 drawing_node_factory: "DrawingNodeFactory"):
        self.ah4lo_lang = ah4lo_lang
        self.oDoc = oDoc
        self.oTextRange = oTextRange
        self.content_node = content_node
        self.drawing_node_factory = drawing_node_factory

        self.oParagraphStyles = self.oDoc.StyleFamilies.ParagraphStyles
        self.oNumberingStyles = self.oDoc.StyleFamilies.NumberingStyles
        self.cur_nodes = []
        self.nodes_stack = [content_node]

    def build(self) -> NodeBuilder:
        cursor_text = self.oTextRange.Text
        for oElement in to_iter(cursor_text):
            oController = self.oDoc.CurrentController

            def action(oController: UnoController = oController,
                       oElement=oElement):
                oController.ViewCursor.gotoRange(oElement.Anchor.Start, False)

            if oElement.supportsService(TEXT_TABLE_SERVICE_NAME):
                self._logger.debug("Table %s", repr(oElement))
                table_name = oElement.Name
                columns_count = oElement.Columns.Count
                rows_count = oElement.Rows.Count
                value = self.ah4lo_lang.writer_table(table_name, columns_count,
                                                     rows_count)
                table_node = NodeBuilder(value, action)
                self.cur_nodes.append(table_node)
            else:
                outline_level = self._get_outline_level(oElement)
                if outline_level > 0:
                    self._flush_nodes()

                    value = self.ah4lo_lang.writer_title(
                        oElement.ListLabelString, oElement.String)
                    title_node = NodeBuilder(value, action)
                    if outline_level < len(self.nodes_stack):
                        self.nodes_stack = self.nodes_stack[:outline_level]
                    self.nodes_stack[-1].append_child(title_node)
                    self.nodes_stack.append(title_node)
                else:
                    par_text = self._shorten(oElement.String, 50)
                    par_node = NodeBuilder(
                        self.ah4lo_lang.paragraph(par_text),
                        action)
                    self.cur_nodes.append(par_node)

                dnf = self.drawing_node_factory
                for oDrawing in dnf.find_drawings(
                        oElement.TextParagraph):
                    drawing_node = dnf.create_drawing_node(
                        oDrawing, action)
                    self.cur_nodes.append(drawing_node)

        self._flush_nodes()
        return self.content_node

    def _shorten(self, text, max_len):
        if len(text) < max_len:
            par_text = text
        else:
            par_text = text[:max_len - 3] + "..."
        return par_text

    def _flush_nodes(self):
        parent_node = self.nodes_stack[-1]
        step = 10
        for i in range(0, len(self.cur_nodes), step):
            nodes = self.cur_nodes[i: i + step]
            pars_node = NodeBuilder(
                self.ah4lo_lang.paragraphs(i + 1, i + len(nodes)))
            for node in nodes:
                pars_node.append_child(node)

            parent_node.append_child(pars_node)
        self.cur_nodes = []

    def _get_outline_level(self, oElement) -> int:
        oStyle = self.oParagraphStyles.getByName(oElement.ParaStyleName)
        while True:
            if oStyle.NumberingStyleName == "":
                return -1
            oNumberingStyle = self.oNumberingStyles.getByName(
                oStyle.NumberingStyleName)
            oRules = oNumberingStyle.NumberingRules

            if oRules and oRules.NumberingIsOutline:
                return oElement.NumberingLevel + 1  # 1 for heading 1
            if oStyle.ParentStyle == "":
                return -1

            oStyle = self.oParagraphStyles.getByName(
                oStyle.ParentStyle)
