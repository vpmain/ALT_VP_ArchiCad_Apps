from fpdf import FPDF
from datetime import date
import math

font_SegoeUI = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\font\segoeui.ttf"
font_SegoeUI_bold = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\font\segoeuib.ttf"
font_SegoeUI_italic = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\font\segoeuii.ttf"
font_SegoeUI_bolditalic = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\font\segoeuiz.ttf"

class PDF(FPDF):
    def __init__(self, form, ori, cols, proj, **kwargs):
        super(PDF, self).__init__(**kwargs)
        self.add_font('Segoe UI', '', font_SegoeUI, uni=True)
        self.add_font('Segoe UI', 'B', font_SegoeUI_bold, uni=True)
        self.add_font('Segoe UI', 'I', font_SegoeUI_italic, uni=True)
        self.add_font('Segoe UI', 'BI', font_SegoeUI_bolditalic, uni=True)

        if form == "Format A4" and ori == "Hochformat":
            self.width = 210
            self.heigth = 297
        elif form == "Format A4" and ori == "Querformat":
            self.width = 297
            self.heigth = 210
        elif form == "Format A3" and ori == "Hochformat":
            self.width = 297
            self.heigth = 420
        elif form == "Format A3" and ori == "Querformat":
            self.width = 420
            self.heigth = 297

        if cols == "1 Spalte":
            self.col = 2
        if cols == "2 Spalten":
            self.col = 4
        if cols == "3 Spalten":
            self.col = 6
        if cols == "4 Spalten":
            self.col = 8
        if cols == "5 Spalten":
            self.col = 10

        self.marginLeft = 15
        self.marginRight = 15
        self.marginTop = 25

        self.referenzWidth = self.width - 10 - 10
        self.singleCellWidth = self.referenzWidth / self.col

        self.newLineMarginVerySmall = 4
        self.newLineMarginSmall = 6
        self.newLineMarginBig = 8
        self.newLineMarginVeryBig = 10

        self.fontStyle = "Segoe UI"
        self.text_color_gray = 140
        self.text_color_black = 0

        self.date = date.today()
        self.proj = proj

        self.set_auto_page_break(auto=True, margin=25)
        self.logopath = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\Template_PDF\Logo\VP_logo.png"

    def header(self):
        # Logo
        self.image(self.logopath, 10, 10, h=12)
        self.set_font(self.fontStyle, 'B', 12)
        self.ln(8)
        self.cell(self.referenzWidth / 2, 5, "", border=0, align='R')
        self.cell(self.referenzWidth / 2, 5, "Raumbuch: " + self.proj, border=0, align='R')
        self.ln(8)
        self.cell(self.referenzWidth, 5, "", border='T', align='R')
        self.ln(10)

    def footer(self):
        # Set position of the footer
        self.set_y(-20)
        # set font
        self.set_font(self.fontStyle, '', 10)
        self.cell(self.referenzWidth, 5, "", border='B', align='R')
        self.ln(5)
        # Page number
        self.cell(self.referenzWidth/2, 10, f"Ausgabe am: {self.date}", border=0, align='L')
        self.cell(self.referenzWidth/2, 10, f"Seite {self.page_no()} von {self.alias_nb_pages()}", border=0, align='R')

    def chapter(self, header):
        self.set_font(self.fontStyle, 'B', 10)
        self.cell(2, 10, "", border=0, align='L')
        self.cell(2 * self.singleCellWidth-4, 5, header, border='B', align='L')
        self.cell(2, 10, "", border=0, align='L')

    def data(self, bez, data):
        self.set_font(self.fontStyle, 'I', 10)
        self.cell(2, 10, "", border=0, align='L')
        self.cell(self.singleCellWidth - 2, 5, bez, border=0, align='L')
        self.set_font(self.fontStyle, '', 10)
        self.cell(self.singleCellWidth - 2, 5, data, border=0, align='L')
        self.cell(2, 10, "", border=0, align='L')

    def empty(self):
        self.set_font(self.fontStyle, 'B', 10)
        self.cell(2, 10, "", border=0, align='L')
        self.cell(2 * self.singleCellWidth-4, 5, "", border=0, align='L')
        self.cell(2, 10, "", border=0, align='L')