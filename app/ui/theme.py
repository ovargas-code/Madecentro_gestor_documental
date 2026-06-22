from __future__ import annotations


PRIMARY = "#F08419"
PRIMARY_HOVER = "#D96F0C"
SECONDARY = "#2D2D2D"
BACKGROUND = "#FFFFFF"
TEXT = "#333333"
MUTED = "#6B7280"
BORDER = "#E5E7EB"
SURFACE = "#F7F7F8"


APP_STYLESHEET = f"""
QMainWindow, QDialog {{
    background: {SURFACE};
    color: {TEXT};
}}
QWidget {{
    color: {TEXT};
    font-family: "Segoe UI";
    font-size: 13px;
}}
QFrame#brandHeader {{
    background: {SECONDARY};
    border: none;
}}
QFrame#brandDivider {{
    background: {PRIMARY};
    border: none;
    border-radius: 1px;
}}
QLabel#brandTitle {{
    color: {BACKGROUND};
    font-size: 21px;
    font-weight: 700;
}}
QLabel#brandSubtitle {{
    color: #D1D5DB;
    font-size: 12px;
}}
QLabel#pageTitle {{
    color: {SECONDARY};
    font-size: 20px;
    font-weight: 700;
}}
QLabel#pageSubtitle, QLabel#mutedText {{
    color: {MUTED};
}}
QLabel#sectionTitle {{
    color: {SECONDARY};
    font-size: 14px;
    font-weight: 700;
}}
QLabel#statusReady {{
    background: #FFF3E8;
    color: #A84C00;
    border: 1px solid #FFD1A6;
    border-radius: 8px;
    padding: 8px 10px;
}}
QLabel#statusMissing {{
    background: #F3F4F6;
    color: {MUTED};
    border: 1px solid {BORDER};
    border-radius: 8px;
    padding: 8px 10px;
}}
QLabel#summaryText {{
    color: {MUTED};
    background: transparent;
    border: none;
    padding: 2px 0;
}}
QLabel#instructionStep {{
    color: {TEXT};
    padding: 3px 4px;
}}
QLabel#instructionWarning {{
    background: #FFF3E8;
    color: #8A3D00;
    border: 1px solid #FFD1A6;
    border-radius: 9px;
    padding: 12px 14px;
    font-weight: 600;
}}
QFrame#card {{
    background: {BACKGROUND};
    border: 1px solid {BORDER};
    border-radius: 12px;
}}
QTabWidget::pane {{
    border: none;
    background: {SURFACE};
    top: -1px;
}}
QTabWidget {{
    background: {SECONDARY};
}}
QTabWidget > QWidget {{
    background: {SURFACE};
}}
QTabBar::tab {{
    background: {SECONDARY};
    color: #D1D5DB;
    border: none;
    padding: 13px 22px;
    min-width: 115px;
}}
QTabBar::tab:selected {{
    background: {PRIMARY};
    color: {BACKGROUND};
    font-weight: 700;
}}
QTabBar::tab:hover:!selected {{
    background: #414141;
    color: {BACKGROUND};
}}
QPushButton {{
    background: {BACKGROUND};
    color: {SECONDARY};
    border: 1px solid #D1D5DB;
    border-radius: 7px;
    padding: 8px 13px;
    min-height: 18px;
    font-weight: 600;
}}
QPushButton:hover {{
    border-color: {PRIMARY};
    color: {PRIMARY};
}}
QPushButton:pressed {{
    background: #FFF3E8;
}}
QPushButton[role="primary"] {{
    background: {PRIMARY};
    color: {BACKGROUND};
    border-color: {PRIMARY};
}}
QPushButton[role="primary"]:hover {{
    background: {PRIMARY_HOVER};
    border-color: {PRIMARY_HOVER};
    color: {BACKGROUND};
}}
QPushButton[role="danger"] {{
    color: #B42318;
    border-color: #FDA29B;
}}
QPushButton[role="danger"]:hover {{
    background: #FEF3F2;
    border-color: #F04438;
}}
QLineEdit, QComboBox {{
    background: {BACKGROUND};
    border: 1px solid #D1D5DB;
    border-radius: 7px;
    padding: 7px 9px;
    min-height: 20px;
    selection-background-color: {PRIMARY};
}}
QLineEdit:focus, QComboBox:focus {{
    border: 2px solid {PRIMARY};
    padding: 6px 8px;
}}
QTableWidget {{
    background: {BACKGROUND};
    alternate-background-color: #FAFAFA;
    border: 1px solid {BORDER};
    border-radius: 8px;
    gridline-color: {BORDER};
    selection-background-color: #FFF0E0;
    selection-color: {TEXT};
}}
QHeaderView::section {{
    background: {SECONDARY};
    color: {BACKGROUND};
    border: none;
    border-right: 1px solid #4A4A4A;
    padding: 9px;
    font-weight: 700;
}}
QTableCornerButton::section {{
    background: {SECONDARY};
    border: none;
}}
QListWidget#templateList {{
    background: {BACKGROUND};
    border: 1px solid {BORDER};
    border-radius: 8px;
    padding: 4px;
    outline: none;
}}
QListWidget#templateList::item {{
    border: 1px solid transparent;
    border-radius: 6px;
    padding: 7px 10px;
    margin: 1px;
}}
QListWidget#templateList::item:hover {{
    background: #FFF7EF;
    border-color: #FFD1A6;
}}
QListWidget#templateList::item:selected {{
    background: #FFF0E0;
    border-color: {PRIMARY};
    color: {SECONDARY};
}}
QScrollBar:vertical {{
    background: #F3F4F6;
    width: 11px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: #C7C9CC;
    border-radius: 5px;
    min-height: 25px;
}}
QScrollBar::handle:vertical:hover {{
    background: {PRIMARY};
}}
QSplitter::handle {{
    background: {BORDER};
}}
"""
