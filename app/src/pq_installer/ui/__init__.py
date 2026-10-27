"""Shared Vector visual system for the desktop Power Query Manager UI.

Presentation only: this module does not touch Excel, DAO or query operations.
The contrast-focused dark palette is the default; light mode remains available.
"""
from string import Template


def vector_colors(dark=True):
    if dark:
        return {
            "window": "#080e12", "surface": "#111b20",
            "surface_alt": "#17252b", "inset": "#0b1418",
            "sidebar": "#0a1216", "border": "#30434a",
            "soft_border": "#23343a", "text": "#ebf4f3",
            "title": "#ffffff", "muted": "#adbec3",
            "muted2": "#95abb2", "accent": "#cff66d",
            "accent_text": "#102016", "accent_hover": "#e1ffa0",
            "teal": "#81e4d3", "selection": "#314c4a",
            "hover": "#23363a", "warning": "#ffd18b",
        }
    return {
        "window": "#f1f6f5", "surface": "#ffffff",
        "surface_alt": "#eaf1ef", "inset": "#f6f9f8",
        "sidebar": "#e7efec", "border": "#b4c8c8",
        "soft_border": "#d4e0de", "text": "#193037",
        "title": "#0e2227", "muted": "#455d64",
        "muted2": "#45636a", "accent": "#3b6a21",
        "accent_text": "#ffffff", "accent_hover": "#275317",
        "teal": "#12645b", "selection": "#cbe7d9",
        "hover": "#dcebe6", "warning": "#875200",
    }


_VECTOR_QSS = Template(r"""
QMainWindow, QDialog, QWidget#vectorCanvas, QWidget#vectorRoot {
    background-color: $window; color: $text;
}
QWidget {
    color: $text; font-family: "Segoe UI", "Inter", sans-serif;
    font-size: 10pt;
}
QLabel, QCheckBox { background: transparent; }
QMainWindow::separator { background: $soft_border; width: 3px; height: 3px; }
QWidget#vectorCanvas, QScrollArea#vectorScroll, QWidget#vectorRoot {
    background: $window;
}
QFrame#vectorSidebar {
    background-color: $sidebar; border-right: 1px solid $border;
}
QFrame#vectorStatusBar, QFrame#vectorFinderBand {
    background-color: $surface; border: 1px solid $border;
    border-radius: 6px;
}
QFrame#vectorMetric {
    background-color: $surface; border: 1px solid $border;
    border-top: 2px solid $accent; border-radius: 5px;
}
QFrame#vectorPanel, QFrame#vectorListPanel {
    background-color: $surface; border: 1px solid $border;
    border-radius: 7px;
}
QLabel#vectorEyebrow {
    color: $accent; font-family: "Consolas", monospace;
    font-size: 9pt; font-weight: 700;
}
QLabel#vectorHero { color: $title; font-size: 23pt; font-weight: 800; }
QLabel#vectorPanelTitle { color: $title; font-size: 12pt; font-weight: 700; }
QLabel#vectorMetricValue {
    color: $title; font-family: "Consolas", monospace;
    font-size: 20pt; font-weight: 700;
}
QLabel#vectorMetricCaption, QLabel#secondaryText, QLabel#vectorHint {
    color: $muted;
}
QLabel#vectorMetricLabel {
    color: $muted; font-family: "Consolas", monospace; font-size: 9pt;
}
QLabel#vectorStatusLabel { color: $teal; font-family: "Consolas", monospace; font-size: 9pt; }
QGroupBox {
    background-color: $surface; color: $text; font-weight: 700;
    border: 1px solid $border; border-radius: 7px;
    margin-top: 15px; padding: 17px 12px 12px 12px;
}
QGroupBox::title {
    subcontrol-origin: margin; subcontrol-position: top left;
    left: 16px; padding: 0px 8px; color: $teal; font-size: 10pt;
}
QPushButton {
    background-color: $surface_alt; color: $text;
    border: 1px solid $border; border-radius: 5px;
    padding: 8px 11px; min-height: 22px; font-weight: 600;
}
QPushButton:hover { background-color: $hover; border-color: $teal; }
QPushButton:pressed { background-color: $selection; }
QPushButton:disabled { color: $muted2; background-color: $surface; border-color: $soft_border; }
QPushButton#primary, QPushButton#vectorPrimary {
    background-color: $accent; color: $accent_text;
    border: 1px solid $accent; font-weight: 800;
}
QPushButton#primary:hover, QPushButton#vectorPrimary:hover {
    background-color: $accent_hover; border-color: $accent_hover;
}
QPushButton#vectorFinderButton {
    background-color: $surface_alt; color: $accent;
    border: 1px solid $border; font-weight: 800;
}
QPushButton#vectorFinderButton:hover {
    background-color: $hover; border-color: $accent;
}
QPushButton#vectorNav {
    background-color: transparent; color: $muted;
    border: 1px solid transparent; border-radius: 7px;
    font-size: 15pt; min-height: 43px; min-width: 43px; padding: 4px;
}
QPushButton#vectorNav:hover { color: $title; background-color: $surface_alt; border-color: $border; }
QPushButton#vectorNavActive {
    background-color: $surface_alt; color: $accent;
    border: 1px solid $accent; border-radius: 7px;
    font-size: 15pt; min-height: 43px; min-width: 43px; padding: 4px;
}
QPushButton#vectorNavActive:hover { background-color: $hover; }
QLineEdit, QComboBox, QPlainTextEdit, QTextEdit, QListWidget,
QTableView, QTableWidget, QSpinBox {
    background-color: $inset; color: $text;
    border: 1px solid $border; border-radius: 5px;
    padding: 7px; selection-background-color: $selection;
    selection-color: $title;
}
QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus,
QListWidget:focus, QTableView:focus, QTableWidget:focus,
QPushButton:focus, QCheckBox:focus {
    border: 2px solid $teal;
}
QComboBox::drop-down { border: 0px; width: 24px; }
QComboBox QAbstractItemView {
    background-color: $surface; color: $text; border: 1px solid $border;
    selection-background-color: $selection;
}
QCheckBox { spacing: 8px; color: $text; }
QCheckBox::indicator { width: 16px; height: 16px; }
QCheckBox::indicator:unchecked {
    background: $surface; border: 1px solid $muted; border-radius: 3px;
}
QHeaderView::section {
    background-color: $surface_alt; color: $muted;
    padding: 9px 7px; border: 0px; border-bottom: 1px solid $border;
    font-size: 9pt; font-weight: 700;
}
QTableView::item, QTableWidget::item { padding: 5px 6px; }
QTableView::item:selected, QTableWidget::item:selected,
QListWidget::item:selected { background: $selection; color: $title; }
QTableView, QTableWidget { gridline-color: $soft_border; alternate-background-color: $surface; }
QScrollBar:vertical { background: $inset; width: 10px; margin: 0px; }
QScrollBar:horizontal { background: $inset; height: 10px; margin: 0px; }
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
    background: $border; border-radius: 4px; min-height: 24px; min-width: 24px;
}
QScrollBar::add-line, QScrollBar::sub-line { height: 0px; width: 0px; }
QDockWidget { color: $title; font-weight: 700; }
QDockWidget::title {
    background: $surface; border-bottom: 1px solid $border;
    text-align: left; padding: 10px 12px;
}
QTabWidget::pane { border: 1px solid $border; background: $surface; }
QTabBar::tab {
    background: $surface; color: $muted; border: 1px solid $border;
    padding: 10px 15px;
}
QTabBar::tab:selected { background: $surface_alt; color: $accent; border-bottom: 2px solid $accent; }
QProgressBar {
    border: 1px solid $border; border-radius: 4px;
    background: $inset; color: $text; text-align: center; min-height: 11px;
}
QProgressBar::chunk { background: $accent; border-radius: 3px; }
QSplitter::handle { background: $soft_border; }
QMenuBar, QMenu, QStatusBar {
    background: $sidebar; color: $muted; border-color: $border;
}
QMenu::item:selected, QMenuBar::item:selected { background: $hover; color: $title; }
QToolTip { background: $surface_alt; color: $text; border: 1px solid $teal; padding: 5px; }
""")


def vector_stylesheet(dark=True):
    """Return a shared style sheet. Bright focus states apply in both themes."""
    return _VECTOR_QSS.substitute(vector_colors(dark))


def apply_vector_theme(widget, dark=True):
    widget.setStyleSheet(vector_stylesheet(dark))
