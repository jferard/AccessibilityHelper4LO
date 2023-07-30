def _with_s(name: str, count: int) -> str:
    if count > 1:
        return "{} {}s".format(count, name)
    else:
        return "{} {}".format(count, name)


class AH4LOLang:
    document_word = "document"
    sheet_word = "sheet"
    column_word = "column"
    row_word = "row"
    used_range_word = "used range"
    masked_word = "masked"
    protected_word = "protected"
    empty_word = "(Empty)"

    @staticmethod
    def from_lang(lang: str) -> "AH4LOLang":
        if lang == "fr":
            return AH4LOLangFr()
        else:
            return AH4LOLangEn()

    def calc_window_title(self, doc_title: str, sheet_count: int) -> str:
        return "{} {}, {}".format(self.document_word.capitalize(),
                                  doc_title,
                                  _with_s(self.sheet_word, sheet_count))

    def sheets(self, count: int) -> str:
        return _with_s(self.sheet_word, count)

    def used_range(self, column_count: int, row_count: int) -> str:
        return "{} : {} × {}".format(
            self.used_range_word.capitalize(),
            _with_s(self.column_word, column_count),
            _with_s(self.row_word, row_count))

    def sheet_description(self, sheet_name: str, is_hidden: bool,
                          is_protected: bool) -> str:
        return self._sheet_description(
            sheet_name, is_hidden, is_protected, self.sheet_word)

    def _sheet_description(self, sheet_name: str, is_hidden: bool,
                           is_protected: bool) -> str:
        comments = []
        if is_hidden:
            comments.append(self.masked_word)
        if is_protected:
            comments.append(self.protected_word)
        if comments:
            name = "{} {} ({})".format(self.sheet_word, sheet_name,
                                       ", ".join(comments))
        else:
            name = "{} {}".format(self.sheet_word, sheet_name)
        return name

    def columns(self, count: int) -> str:
        return _with_s(self.column_word, count).capitalize()

    def get_type_name(self, data_type: int) -> str:
        return {
            0: "All",
            1: "Defined",
            2: "Date",
            4: "Time",
            8: "Currency",
            16: "Number",
            32: "Scientific",
            64: "Fraction",
            128: "Percent",
            256: "Text",
            6: "Datetime",
            1024: "Logical",
            2048: "Undefined",
            4096: "Empty",
            8196: "Duration",
        }.get(data_type, "All")


class AH4LOLangEn(AH4LOLang):
    pass


class AH4LOLangFr(AH4LOLang):
    document_word = "document"
    sheet_word = "feuille"
    column_word = "colonne"
    row_word = "ligne"
    used_range_word = "plage utilisée"
    empty_word = "(Vide)"
    masked_word = "masquée"
    protected_word = "protégée"

    def get_type_name(self, data_type: int) -> str:
        return {
            0: "Tout",
            1: "Défini",
            2: "Date",
            4: "Temps",
            8: "Monétaire",
            16: "Nombre",
            32: "Scientifique",
            64: "Fraction",
            128: "Pourcentage",
            256: "Texte",
            6: "Date-temps",
            1024: "Logique",
            2048: "Indéfini",
            4096: "Vide",
            8196: "Durée",
        }.get(data_type, "Tout")
