import logging
from typing import Optional, cast, List

from ah4lo_lang import AH4LOLang
from ah4lo_tree import Node, NodeBuilder
from lo_helper import guess_format_id, get_type_id, extract_values
from py4lo_helper import get_used_range, to_iter
from py4lo_typing import UnoSpreadsheet, UnoRange, UnoSheet

BASE_FRAME_SERVICE_NAME = "com.sun.star.text.BaseFrame"

TEXT_TABLE_SERVICE_NAME = "com.sun.star.text.TextTable"

PARAGRAPH_SERVICE_NAME = "com.sun.star.text.Paragraph"


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


class WriterDocumentNodeFactory:
    _logger = logging.getLogger(__name__)

    def __init__(self, ah4lo_lang: AH4LOLang, oDoc: UnoSpreadsheet):
        self.ah4lo_lang = ah4lo_lang
        self.oDoc = oDoc
        self.oParagraphStyles = self.oDoc.StyleFamilies.ParagraphStyles
        self.oNumberingStyles = self.oDoc.StyleFamilies.NumberingStyles

    def get_root(self) -> Node:
        oProperties = self.oDoc.DocumentProperties
        root_node = NodeBuilder(oProperties.Title)
        information_node = self._get_informations()
        root_node.append_child(information_node)
        content_node = self._get_content()
        root_node.append_child(content_node)
        root_node.freeze_as_root()
        return root_node

    def _get_content(self) -> NodeBuilder:
        content_node = NodeBuilder(self.ah4lo_lang.content())
        oCursor = self.oDoc.Text.createTextCursor()
        nodes = [content_node]
        oCursor.gotoStart(False)
        oCursor.gotoEnd(True)

        oGraphicObjects = self.oDoc.GraphicObjects
        self._logger.debug("%s", repr(oGraphicObjects))
        graphic_objects_by_paragraph = {
            oGo.Anchor.TextParagraph: oGo for oGo in to_iter(oGraphicObjects)
        }

        oTextFrames = self.oDoc.TextFrames
        self._logger.debug("%s", repr(oTextFrames))
        text_frames_by_paragraph = {
            oTf.Anchor.TextParagraph: oTf for oTf in to_iter(oTextFrames)
        }

        cur_nodes = []
        for oElement in to_iter(oCursor.Text):
            oController = self.oDoc.CurrentController

            if not oElement.supportsService(PARAGRAPH_SERVICE_NAME):
                self._logger.debug("Other %s",
                                   repr(oElement))

            def action(oController=oController, oElement=oElement):
                oController.ViewCursor.gotoRange(oElement.Anchor.Start, False)

            if oElement.supportsService(TEXT_TABLE_SERVICE_NAME):
                self._logger.debug("Table %s", repr(oElement))
                table_name = oElement.Name
                columns_count = oElement.Columns.Count
                rows_count = oElement.Rows.Count
                value = self.ah4lo_lang.writer_table(table_name, columns_count,
                                                     rows_count)
                table_node = NodeBuilder(value, action)
                cur_nodes.append(table_node)
            else:
                outline_level = self._get_outline_level(oElement)
                if outline_level > 0:
                    self._flush_nodes(nodes[-1], cur_nodes)

                    value = self.ah4lo_lang.writer_title(
                        oElement.ListLabelString, oElement.String)
                    title_node = NodeBuilder(value, action)
                    if outline_level < len(nodes):
                        nodes = nodes[:outline_level]
                    nodes[-1].append_child(title_node)
                    nodes.append(title_node)
                else:
                    par_text = oElement.String[:50] + "..."
                    par_node = NodeBuilder(
                        self.ah4lo_lang.paragraph(par_text),
                        action)
                    cur_nodes.append(par_node)

                try:
                    oGo = graphic_objects_by_paragraph[oElement.TextParagraph]
                except KeyError:
                    pass
                else:
                    go_node = NodeBuilder(
                        self.ah4lo_lang.graphic_object(oGo.Name), action)
                    cur_nodes.append(go_node)

                try:
                    oTf = text_frames_by_paragraph[oElement.TextParagraph]
                except KeyError:
                    pass
                else:
                    tf_node = NodeBuilder(
                        self.ah4lo_lang.text_frame(oTf.Name), action)
                    for oTfElement in to_iter(oTf):
                        # TODO: look for titles, paragraphs
                        try:
                            oGo = graphic_objects_by_paragraph[
                                oTfElement.TextParagraph]
                        except KeyError:
                            pass
                        else:
                            go_node = NodeBuilder(
                                self.ah4lo_lang.graphic_object(oGo.Name),
                                action)
                            tf_node.append_child(go_node)

                    cur_nodes.append(tf_node)

        self._flush_nodes(nodes[-1], cur_nodes)
        return content_node

    def _flush_nodes(self, parent_node: NodeBuilder,
                     cur_nodes: List[NodeBuilder]):
        step = 10
        self._flush_nodes_aux(parent_node, cur_nodes, step)
        cur_nodes.clear()

    def _flush_nodes_aux(self, parent_node, cur_nodes, step):
        for i in range(0, len(cur_nodes), step):
            nodes = cur_nodes[i: i + step]
            pars_node = NodeBuilder(
                self.ah4lo_lang.paragraphs(i + 1, i + len(nodes)))
            for node in nodes:
                pars_node.append_child(node)

            parent_node.append_child(pars_node)

    def _get_informations(self) -> NodeBuilder:
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
