import sys
import pandas as pd
import matplotlib.ticker as ticker

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QLabel,
    QFormLayout, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QComboBox, QDateEdit, QGroupBox, QGridLayout,
    QScrollArea, QSizePolicy, QHeaderView, QAbstractScrollArea,
    QHBoxLayout, QGraphicsOpacityEffect, QFileDialog
)
from PyQt5.QtCore import QDate, Qt, QPropertyAnimation, QEasingCurve

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from src.analysis_service import run_full_analysis
from src.persistence import save_project, load_project


class AnalyticsWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Business Intelligence Dashboard")
        self.setGeometry(200, 100, 1400, 900)

        layout = QVBoxLayout()

        self.kpi_revenue = QLabel("₦0")
        self.kpi_orders = QLabel("0")
        self.kpi_units = QLabel("0")
        self.kpi_cost = QLabel("N/A")
        self.kpi_profit = QLabel("N/A")
        self.kpi_margin = QLabel("N/A")

        self.kpi_labels = [
            self.kpi_revenue, self.kpi_orders, self.kpi_units,
            self.kpi_cost, self.kpi_profit, self.kpi_margin
        ]

        for lbl in self.kpi_labels:
            lbl.setAlignment(Qt.AlignCenter)

        kpi_layout = QGridLayout()
        kpi_layout.addWidget(QLabel("Total Revenue"), 0, 0)
        kpi_layout.addWidget(QLabel("Total Orders"), 0, 1)
        kpi_layout.addWidget(QLabel("Total Units Sold"), 0, 2)

        kpi_layout.addWidget(self.kpi_revenue, 1, 0)
        kpi_layout.addWidget(self.kpi_orders, 1, 1)
        kpi_layout.addWidget(self.kpi_units, 1, 2)

        kpi_layout.addWidget(QLabel("Total Cost"), 2, 0)
        kpi_layout.addWidget(QLabel("Total Profit"), 2, 1)
        kpi_layout.addWidget(QLabel("Average Margin"), 2, 2)

        kpi_layout.addWidget(self.kpi_cost, 3, 0)
        kpi_layout.addWidget(self.kpi_profit, 3, 1)
        kpi_layout.addWidget(self.kpi_margin, 3, 2)

        self.chart_canvas = FigureCanvas(Figure(figsize=(12, 8)))

        layout.addLayout(kpi_layout)
        layout.addWidget(self.chart_canvas)

        self.setLayout(layout)

    def apply_theme(self, theme):
        if theme == "light":
            card_style = """
                QLabel {
                    background:#F5F7FA;
                    border:1px solid #C9D2DC;
                    border-radius:10px;
                    padding:15px;
                    font-size:20px;
                    font-weight:bold;
                    color:#0F6CBD;
                }
            """
            figure_bg = "#FFFFFF"
        else:
            card_style = """
                QLabel {
                    background:#252525;
                    border:1px solid #3C3C3C;
                    border-radius:10px;
                    padding:15px;
                    font-size:20px;
                    font-weight:bold;
                    color:#57C7FF;
                }
            """
            figure_bg = "#151515"

        for lbl in self.kpi_labels:
            lbl.setStyleSheet(card_style)

        self.chart_canvas.figure.patch.set_facecolor(figure_bg)
        self.chart_canvas.draw_idle()


class BusinessAnalyticsApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Business Analytics System")
        self.setGeometry(100, 100, 1450, 920)

        self.products = []
        self.customers = []
        self.sales = []
        self.kpi_animations = []
        self.is_refreshing_tables = False
        self.current_theme = "dark"

        self.tabs = QTabWidget()

        self.setup_tab = QWidget()
        self.sales_tab = QWidget()
        self.report_tab = QWidget()

        self.tabs.addTab(self.setup_tab, "Business Setup")
        self.tabs.addTab(self.sales_tab, "Sales Entry")
        self.tabs.addTab(self.report_tab, "Analytics Summary")

        self.setup_ui()
        self.analytics_window = AnalyticsWindow()
        self.build_main_layout()

        self.apply_theme("dark")

    def create_scrollable_tab(self, content_widget):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(content_widget)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        wrapper = QVBoxLayout()
        wrapper.addWidget(scroll)

        container = QWidget()
        container.setLayout(wrapper)
        return container

    def configure_table(self, table, min_height=220):
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        table.setMinimumHeight(min_height)
        table.setAlternatingRowColors(True)
        table.setWordWrap(False)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.verticalHeader().setVisible(True)
        table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)

    def make_collapsible(self, group_box, checked=True):
        group_box.setCheckable(True)
        group_box.setChecked(checked)

        def toggle_group(state):
            for i in range(group_box.layout().count()):
                item = group_box.layout().itemAt(i)
                widget = item.widget()
                if widget is not None:
                    widget.setVisible(state)

        group_box.toggled.connect(toggle_group)
        toggle_group(checked)

    def build_main_layout(self):
        self.tabs.tabBar().hide()

        main_container = QWidget()
        main_layout = QHBoxLayout(main_container)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        sidebar = QGroupBox("Navigation")
        sidebar_layout = QVBoxLayout()

        self.btn_setup = QPushButton("Business Setup")
        self.btn_sales = QPushButton("Sales Entry")
        self.btn_analytics = QPushButton("Analytics Summary")

        self.btn_setup.setObjectName("navButton")
        self.btn_sales.setObjectName("navButton")
        self.btn_analytics.setObjectName("navButton")

        self.btn_setup.clicked.connect(lambda: self.tabs.setCurrentIndex(0))
        self.btn_sales.clicked.connect(lambda: self.tabs.setCurrentIndex(1))
        self.btn_analytics.clicked.connect(lambda: self.tabs.setCurrentIndex(2))

        sidebar_layout.addWidget(self.btn_setup)
        sidebar_layout.addWidget(self.btn_sales)
        sidebar_layout.addWidget(self.btn_analytics)
        sidebar_layout.addStretch()

        sidebar.setLayout(sidebar_layout)
        sidebar.setFixedWidth(230)

        main_layout.addWidget(sidebar)
        main_layout.addWidget(self.tabs)

        self.setCentralWidget(main_container)

    def setup_ui(self):
        # ================= BUSINESS SETUP TAB CONTENT =================
        setup_content = QWidget()
        main_layout = QVBoxLayout(setup_content)
        main_layout.setSpacing(16)

        company_group = QGroupBox("Company Information")
        company_layout = QFormLayout()

        self.company_input = QLineEdit()

        self.theme_selector = QComboBox()
        self.theme_selector.addItems(["Dark Mode", "Light Mode"])
        self.theme_selector.currentTextChanged.connect(self.on_theme_changed)

        company_layout.addRow("Company Name:", self.company_input)
        company_layout.addRow("Theme:", self.theme_selector)

        company_group.setLayout(company_layout)
        self.make_collapsible(company_group, checked=True)

        product_group = QGroupBox("Product Management")
        product_layout = QGridLayout()

        self.pid = QLineEdit()
        self.pname = QLineEdit()
        self.pcat = QComboBox()
        self.pcat.addItems(["Laptop", "Mobile", "Accessory", "Networking", "Software", "Other"])
        self.pprice = QLineEdit()
        self.pcost = QLineEdit()
        self.pcost.setPlaceholderText("Optional")

        product_layout.addWidget(QLabel("Product ID"), 0, 0)
        product_layout.addWidget(self.pid, 0, 1)
        product_layout.addWidget(QLabel("Product Name"), 1, 0)
        product_layout.addWidget(self.pname, 1, 1)
        product_layout.addWidget(QLabel("Category"), 2, 0)
        product_layout.addWidget(self.pcat, 2, 1)
        product_layout.addWidget(QLabel("Unit Price"), 3, 0)
        product_layout.addWidget(self.pprice, 3, 1)
        product_layout.addWidget(QLabel("Cost Price (Optional)"), 4, 0)
        product_layout.addWidget(self.pcost, 4, 1)

        self.add_product_btn = QPushButton("Add Product")
        self.delete_product_btn = QPushButton("Delete Selected Product")

        product_layout.addWidget(self.add_product_btn, 5, 0)
        product_layout.addWidget(self.delete_product_btn, 5, 1)

        self.product_search = QLineEdit()
        self.product_search.setPlaceholderText("Search products...")
        product_layout.addWidget(self.product_search, 6, 0, 1, 2)

        self.product_table = QTableWidget()
        self.product_table.setColumnCount(5)
        self.product_table.setHorizontalHeaderLabels(["ID", "Name", "Category", "Unit Price", "Cost Price"])
        self.configure_table(self.product_table, min_height=240)
        product_layout.addWidget(self.product_table, 7, 0, 1, 2)

        product_group.setLayout(product_layout)
        self.make_collapsible(product_group, checked=True)

        customer_group = QGroupBox("Customer Management")
        customer_layout = QGridLayout()

        self.cid = QLineEdit()
        self.cname = QLineEdit()
        self.ccity = QLineEdit()
        self.cregion = QComboBox()
        self.cregion.addItems([
            "North West", "North East", "North Central",
            "South West", "South East", "South South"
        ])

        customer_layout.addWidget(QLabel("Customer ID"), 0, 0)
        customer_layout.addWidget(self.cid, 0, 1)
        customer_layout.addWidget(QLabel("Customer Name"), 1, 0)
        customer_layout.addWidget(self.cname, 1, 1)
        customer_layout.addWidget(QLabel("City"), 2, 0)
        customer_layout.addWidget(self.ccity, 2, 1)
        customer_layout.addWidget(QLabel("Region"), 3, 0)
        customer_layout.addWidget(self.cregion, 3, 1)

        self.add_customer_btn = QPushButton("Add Customer")
        self.delete_customer_btn = QPushButton("Delete Selected Customer")

        customer_layout.addWidget(self.add_customer_btn, 4, 0)
        customer_layout.addWidget(self.delete_customer_btn, 4, 1)

        self.customer_search = QLineEdit()
        self.customer_search.setPlaceholderText("Search customers...")
        customer_layout.addWidget(self.customer_search, 5, 0, 1, 2)

        self.customer_table = QTableWidget()
        self.customer_table.setColumnCount(4)
        self.customer_table.setHorizontalHeaderLabels(["ID", "Name", "City", "Region"])
        self.configure_table(self.customer_table, min_height=240)
        customer_layout.addWidget(self.customer_table, 6, 0, 1, 2)

        customer_group.setLayout(customer_layout)
        self.make_collapsible(customer_group, checked=False)

        main_layout.addWidget(company_group)
        main_layout.addWidget(product_group)
        main_layout.addWidget(customer_group)
        main_layout.addStretch()

        self.setup_tab.setLayout(QVBoxLayout())
        self.setup_tab.layout().addWidget(self.create_scrollable_tab(setup_content))

        # ================= SALES TAB CONTENT =================
        sales_content = QWidget()
        sales_layout = QVBoxLayout(sales_content)
        sales_layout.setSpacing(16)

        sales_group = QGroupBox("Sales Entry")
        sales_form = QFormLayout()

        self.oid = QLineEdit()
        self.sdate = QDateEdit()
        self.sdate.setCalendarPopup(True)
        self.sdate.setDate(QDate.currentDate())

        self.scid = QComboBox()
        self.spid = QComboBox()
        self.sqty = QLineEdit()

        sales_form.addRow("Order ID:", self.oid)
        sales_form.addRow("Date:", self.sdate)
        sales_form.addRow("Customer:", self.scid)
        sales_form.addRow("Product:", self.spid)
        sales_form.addRow("Quantity:", self.sqty)

        sales_group.setLayout(sales_form)
        self.make_collapsible(sales_group, checked=True)

        self.add_sale_btn = QPushButton("Add Sale")
        self.delete_sale_btn = QPushButton("Delete Selected Sale")

        button_row = QHBoxLayout()
        button_row.addWidget(self.add_sale_btn)
        button_row.addWidget(self.delete_sale_btn)

        self.sales_search = QLineEdit()
        self.sales_search.setPlaceholderText("Search sales...")

        self.sales_table = QTableWidget()
        self.sales_table.setColumnCount(5)
        self.sales_table.setHorizontalHeaderLabels(["Order ID", "Date", "Customer", "Product", "Qty"])
        self.configure_table(self.sales_table, min_height=420)

        sales_layout.addWidget(sales_group)
        sales_layout.addLayout(button_row)
        sales_layout.addWidget(self.sales_search)
        sales_layout.addWidget(self.sales_table)
        sales_layout.addStretch()

        self.sales_tab.setLayout(QVBoxLayout())
        self.sales_tab.layout().addWidget(self.create_scrollable_tab(sales_content))

        # ================= ANALYTICS SUMMARY TAB CONTENT =================
        report_content = QWidget()
        report_layout = QVBoxLayout(report_content)
        report_layout.setSpacing(16)

        self.analytics_title = QLabel("Business Intelligence Summary")
        self.analytics_title.setStyleSheet("font-size: 18px; font-weight: bold;")

        self.company_preview = QLabel("Company: Not set")
        self.dataset_info = QLabel(
            "Dataset Status:\n"
            "Products: 0\n"
            "Customers: 0\n"
            "Sales: 0"
        )

        self.generate_btn = QPushButton("RUN FULL BUSINESS ANALYSIS")
        self.generate_btn.setObjectName("primaryButton")

        self.open_dashboard_btn = QPushButton("OPEN BI DASHBOARD")
        self.export_excel_btn = QPushButton("Export Excel Report")
        self.export_png_btn = QPushButton("Export Dashboard PNG")
        self.save_btn = QPushButton("Save Project")
        self.load_btn = QPushButton("Load Project")

        self.status_output = QLabel("System ready.")
        self.status_output.setStyleSheet("color: lightgreen;")

        kpi_group = QGroupBox("Quick KPI Summary")
        kpi_layout = QGridLayout()

        self.kpi_revenue = QLabel("₦0")
        self.kpi_orders = QLabel("0")
        self.kpi_units = QLabel("0")
        self.kpi_cost = QLabel("N/A")
        self.kpi_profit = QLabel("N/A")
        self.kpi_margin = QLabel("N/A")

        self.main_kpi_labels = [
            self.kpi_revenue, self.kpi_orders, self.kpi_units,
            self.kpi_cost, self.kpi_profit, self.kpi_margin
        ]

        for lbl in self.main_kpi_labels:
            lbl.setAlignment(Qt.AlignCenter)

        kpi_layout.addWidget(QLabel("Total Revenue"), 0, 0)
        kpi_layout.addWidget(QLabel("Total Orders"), 0, 1)
        kpi_layout.addWidget(QLabel("Total Units Sold"), 0, 2)
        kpi_layout.addWidget(QLabel("Total Cost"), 2, 0)
        kpi_layout.addWidget(QLabel("Total Profit"), 2, 1)
        kpi_layout.addWidget(QLabel("Average Margin"), 2, 2)

        kpi_layout.addWidget(self.kpi_revenue, 1, 0)
        kpi_layout.addWidget(self.kpi_orders, 1, 1)
        kpi_layout.addWidget(self.kpi_units, 1, 2)
        kpi_layout.addWidget(self.kpi_cost, 3, 0)
        kpi_layout.addWidget(self.kpi_profit, 3, 1)
        kpi_layout.addWidget(self.kpi_margin, 3, 2)

        kpi_group.setLayout(kpi_layout)
        self.make_collapsible(kpi_group, checked=True)

        report_layout.addWidget(self.analytics_title)
        report_layout.addWidget(self.company_preview)
        report_layout.addWidget(self.dataset_info)
        report_layout.addWidget(self.generate_btn)
        report_layout.addWidget(self.open_dashboard_btn)
        report_layout.addWidget(self.export_excel_btn)
        report_layout.addWidget(self.export_png_btn)
        report_layout.addWidget(self.save_btn)
        report_layout.addWidget(self.load_btn)
        report_layout.addWidget(self.status_output)
        report_layout.addWidget(kpi_group)
        report_layout.addStretch()

        self.report_tab.setLayout(QVBoxLayout())
        self.report_tab.layout().addWidget(self.create_scrollable_tab(report_content))

        self.add_product_btn.clicked.connect(self.add_product)
        self.delete_product_btn.clicked.connect(self.delete_selected_product)

        self.add_customer_btn.clicked.connect(self.add_customer)
        self.delete_customer_btn.clicked.connect(self.delete_selected_customer)

        self.add_sale_btn.clicked.connect(self.add_sale)
        self.delete_sale_btn.clicked.connect(self.delete_selected_sale)

        self.generate_btn.clicked.connect(self.generate_report)
        self.open_dashboard_btn.clicked.connect(self.open_dashboard)
        self.export_excel_btn.clicked.connect(self.export_excel_report)
        self.export_png_btn.clicked.connect(self.export_dashboard_png)
        self.save_btn.clicked.connect(self.save_project_data)
        self.load_btn.clicked.connect(self.load_project_data)

        self.product_table.itemChanged.connect(self.on_product_table_changed)
        self.customer_table.itemChanged.connect(self.on_customer_table_changed)
        self.sales_table.itemChanged.connect(self.on_sales_table_changed)

        self.product_search.textChanged.connect(self.filter_product_table)
        self.customer_search.textChanged.connect(self.filter_customer_table)
        self.sales_search.textChanged.connect(self.filter_sales_table)

    def on_theme_changed(self, theme_text):
        if theme_text == "Light Mode":
            self.apply_theme("light")
        else:
            self.apply_theme("dark")

    def get_chart_palette(self):
        if self.current_theme == "light":
            return {
                "figure_bg": "#FFFFFF",
                "axes_bg": "#FFFFFF",
                "text": "#1F1F1F",
                "muted_text": "#3F3F3F",
                "grid": "#D6DCE5",
                "spine": "#B8C2CC",
                "kpi_bg": "#F5F7FA",
                "kpi_border": "#C9D2DC",
                "kpi_text": "#0F6CBD",
                "product": "#2E75B6",
                "product_edge": "#6FA8DC",
                "region": "#70AD47",
                "region_edge": "#A9D18E",
                "trend": "#FFC000",
                "trend_marker": "#FFD966",
                "customer": "#ED7D31",
                "customer_edge": "#F4B183",
            }
        return {
            "figure_bg": "#111111",
            "axes_bg": "#151515",
            "text": "#F5F5F5",
            "muted_text": "#D9D9D9",
            "grid": "#3A3A3A",
            "spine": "#555555",
            "kpi_bg": "#252525",
            "kpi_border": "#3C3C3C",
            "kpi_text": "#57C7FF",
            "product": "#57C7FF",
            "product_edge": "#8BE9FD",
            "region": "#7ED957",
            "region_edge": "#B6F09C",
            "trend": "#FFD166",
            "trend_marker": "#FFE29A",
            "customer": "#FF8C69",
            "customer_edge": "#FFC2B0",
        }

    def get_stylesheet(self, theme):
        if theme == "light":
            return """
                QMainWindow { background-color:#F4F6F8; }
                QWidget {
                    background-color:#F4F6F8;
                    color:#1F2933;
                    font-family:Segoe UI;
                    font-size:12px;
                }
                QTabWidget::pane { border:0; }
                QTabBar::tab {
                    background:#E5EAF0;
                    color:#1F2933;
                    padding:10px;
                    margin:2px;
                    border-radius:6px;
                }
                QTabBar::tab:selected {
                    background:#D5DEE8;
                    font-weight:bold;
                }
                QLineEdit, QComboBox, QDateEdit {
                    background:#FFFFFF;
                    color:#1F2933;
                    border:1px solid #C7D0D9;
                    padding:6px;
                    border-radius:4px;
                }
                QPushButton {
                    background:#E3E8EE;
                    color:#1F2933;
                    padding:8px;
                    border-radius:6px;
                    font-weight:bold;
                    border:1px solid #C7D0D9;
                }
                QPushButton:hover { background:#D8E0E8; }
                QPushButton#primaryButton {
                    background-color:#0F6CBD;
                    color:white;
                    border:1px solid #0F6CBD;
                    font-size:16px;
                    padding:15px;
                    border-radius:8px;
                }
                QPushButton#primaryButton:hover {
                    background-color:#0B5CAD;
                }
                QPushButton#navButton {
                    text-align:left;
                    padding:12px;
                    font-size:13px;
                    border-radius:8px;
                }
                QPushButton#navButton:hover {
                    background-color:#DCE4EC;
                }
                QTableWidget {
                    background:#FFFFFF;
                    alternate-background-color:#F7F9FB;
                    color:#1F2933;
                    gridline-color:#D6DCE5;
                    selection-background-color:#CFE8FF;
                    selection-color:#1F2933;
                }
                QHeaderView::section {
                    background:#E6EBF1;
                    color:#1F2933;
                    padding:5px;
                    border:none;
                    font-weight:bold;
                }
                QGroupBox {
                    border:1px solid #C7D0D9;
                    border-radius:6px;
                    margin-top:10px;
                    font-weight:bold;
                    padding:10px;
                    background-color:#F9FBFC;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left:10px;
                    padding:0 5px 0 5px;
                }
                QGroupBox::indicator {
                    width:14px;
                    height:14px;
                }
                QScrollArea {
                    border:none;
                }
            """
        return """
            QMainWindow { background-color:#1E1E1E; }
            QWidget {
                background-color:#1E1E1E;
                color:#EAEAEA;
                font-family:Segoe UI;
                font-size:12px;
            }
            QTabWidget::pane { border:0; }
            QTabBar::tab {
                background:#2A2A2A;
                padding:10px;
                margin:2px;
                border-radius:6px;
            }
            QTabBar::tab:selected {
                background:#3A3A3A;
                font-weight:bold;
            }
            QLineEdit, QComboBox, QDateEdit {
                background:#2A2A2A;
                color:#EAEAEA;
                border:1px solid #444;
                padding:6px;
                border-radius:4px;
            }
            QPushButton {
                background:#3A3A3A;
                color:#EAEAEA;
                padding:8px;
                border-radius:6px;
                font-weight:bold;
                border:1px solid #444;
            }
            QPushButton:hover { background:#4A4A4A; }
            QPushButton#primaryButton {
                background-color:#0078D7;
                color:white;
                font-size:16px;
                padding:15px;
                border-radius:8px;
                border:1px solid #0078D7;
            }
            QPushButton#primaryButton:hover {
                background-color:#1890FF;
            }
            QPushButton#navButton {
                text-align:left;
                padding:12px;
                font-size:13px;
                border-radius:8px;
            }
            QPushButton#navButton:hover {
                background-color:#505050;
            }
            QTableWidget {
                background:#252525;
                alternate-background-color:#202020;
                color:#EAEAEA;
                gridline-color:#444;
                selection-background-color:#0F6CBD;
                selection-color:white;
            }
            QHeaderView::section {
                background:#333;
                color:#EAEAEA;
                padding:5px;
                border:none;
                font-weight:bold;
            }
            QGroupBox {
                border:1px solid #444;
                border-radius:6px;
                margin-top:10px;
                font-weight:bold;
                padding:10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left:10px;
                padding:0 5px 0 5px;
            }
            QGroupBox::indicator {
                width:14px;
                height:14px;
            }
            QScrollArea {
                border:none;
            }
        """

    def apply_kpi_styles(self):
        palette = self.get_chart_palette()
        style = f"""
            QLabel {{
                background-color: {palette['kpi_bg']};
                border: 1px solid {palette['kpi_border']};
                border-radius: 10px;
                padding: 12px;
                font-size: 18px;
                font-weight: bold;
                color: {palette['kpi_text']};
            }}
        """
        for lbl in self.main_kpi_labels:
            lbl.setStyleSheet(style)

        self.analytics_window.apply_theme(self.current_theme)

        if self.current_theme == "light":
            self.status_output.setStyleSheet("color: #1E7E34;")
        else:
            self.status_output.setStyleSheet("color: lightgreen;")

    def apply_theme(self, theme):
        self.current_theme = theme
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet(self.get_stylesheet(theme))

        self.apply_kpi_styles()
        self.render_charts()

    def animate_kpi(self, label):
        effect = QGraphicsOpacityEffect(label)
        label.setGraphicsEffect(effect)

        animation = QPropertyAnimation(effect, b"opacity")
        animation.setDuration(600)
        animation.setStartValue(0.35)
        animation.setEndValue(1.0)
        animation.setEasingCurve(QEasingCurve.InOutQuad)
        animation.start()

        self.kpi_animations.append(animation)

    def update_status(self):
        self.company_preview.setText(f"Company: {self.company_input.text() or 'Not set'}")
        self.dataset_info.setText(
            f"Dataset Status:\n"
            f"Products: {len(self.products)}\n"
            f"Customers: {len(self.customers)}\n"
            f"Sales: {len(self.sales)}"
        )

    def rebuild_dropdowns(self):
        self.spid.clear()
        self.scid.clear()

        for p in self.products:
            self.spid.addItem(p["Product_ID"])

        for c in self.customers:
            self.scid.addItem(c["Customer_ID"])

    def filter_table(self, table, search_text):
        text = search_text.strip().lower()

        for row in range(table.rowCount()):
            row_match = False
            for col in range(table.columnCount()):
                item = table.item(row, col)
                if item and text in item.text().lower():
                    row_match = True
                    break
            table.setRowHidden(row, not row_match if text else False)

    def filter_product_table(self):
        self.filter_table(self.product_table, self.product_search.text())

    def filter_customer_table(self):
        self.filter_table(self.customer_table, self.customer_search.text())

    def filter_sales_table(self):
        self.filter_table(self.sales_table, self.sales_search.text())

    def add_product(self):
        try:
            price = float(self.pprice.text())
            if price <= 0:
                QMessageBox.warning(self, "Error", "Unit Price must be greater than zero")
                return
        except Exception:
            QMessageBox.warning(self, "Error", "Invalid Unit Price format")
            return

        cost_price_text = self.pcost.text().strip()
        cost_price = None

        if cost_price_text:
            try:
                cost_price = float(cost_price_text)
                if cost_price < 0:
                    QMessageBox.warning(self, "Error", "Cost Price cannot be negative")
                    return
            except Exception:
                QMessageBox.warning(self, "Error", "Invalid Cost Price format")
                return

        product_id = self.pid.text().strip()
        if not product_id:
            QMessageBox.warning(self, "Error", "Product ID is required")
            return

        for p in self.products:
            if p["Product_ID"] == product_id:
                QMessageBox.warning(self, "Error", "Product ID already exists")
                return

        product = {
            "Product_ID": product_id,
            "Product_Name": self.pname.text().strip(),
            "Category": self.pcat.currentText(),
            "Unit_Price": price,
            "Cost_Price": cost_price
        }

        self.products.append(product)
        self.rebuild_dropdowns()

        self.is_refreshing_tables = True
        row = self.product_table.rowCount()
        self.product_table.insertRow(row)
        self.product_table.setItem(row, 0, QTableWidgetItem(product["Product_ID"]))
        self.product_table.setItem(row, 1, QTableWidgetItem(product["Product_Name"]))
        self.product_table.setItem(row, 2, QTableWidgetItem(product["Category"]))
        self.product_table.setItem(row, 3, QTableWidgetItem(str(product["Unit_Price"])))
        self.product_table.setItem(
            row, 4,
            QTableWidgetItem("" if product["Cost_Price"] is None else str(product["Cost_Price"]))
        )
        self.is_refreshing_tables = False

        self.update_status()
        self.filter_product_table()

    def add_customer(self):
        customer_id = self.cid.text().strip()
        if not customer_id:
            QMessageBox.warning(self, "Error", "Customer ID is required")
            return

        for c in self.customers:
            if c["Customer_ID"] == customer_id:
                QMessageBox.warning(self, "Error", "Customer ID already exists")
                return

        customer = {
            "Customer_ID": customer_id,
            "Customer_Name": self.cname.text().strip(),
            "City": self.ccity.text().strip(),
            "Region": self.cregion.currentText()
        }

        self.customers.append(customer)
        self.rebuild_dropdowns()

        self.is_refreshing_tables = True
        row = self.customer_table.rowCount()
        self.customer_table.insertRow(row)
        self.customer_table.setItem(row, 0, QTableWidgetItem(customer["Customer_ID"]))
        self.customer_table.setItem(row, 1, QTableWidgetItem(customer["Customer_Name"]))
        self.customer_table.setItem(row, 2, QTableWidgetItem(customer["City"]))
        self.customer_table.setItem(row, 3, QTableWidgetItem(customer["Region"]))
        self.is_refreshing_tables = False

        self.update_status()
        self.filter_customer_table()

    def add_sale(self):
        try:
            qty = int(self.sqty.text())
            if qty <= 0:
                QMessageBox.warning(self, "Error", "Quantity must be greater than zero")
                return
        except Exception:
            QMessageBox.warning(self, "Error", "Invalid quantity format")
            return

        order_id = self.oid.text().strip()
        if not order_id:
            QMessageBox.warning(self, "Error", "Order ID is required")
            return

        for s in self.sales:
            if s["Order_ID"] == order_id:
                QMessageBox.warning(self, "Error", "Order ID already exists")
                return

        sale = {
            "Order_ID": order_id,
            "Date": self.sdate.date().toString("yyyy-MM-dd"),
            "Customer_ID": self.scid.currentText(),
            "Product_ID": self.spid.currentText(),
            "Quantity": qty
        }

        self.sales.append(sale)

        self.is_refreshing_tables = True
        row = self.sales_table.rowCount()
        self.sales_table.insertRow(row)
        self.sales_table.setItem(row, 0, QTableWidgetItem(sale["Order_ID"]))
        self.sales_table.setItem(row, 1, QTableWidgetItem(sale["Date"]))
        self.sales_table.setItem(row, 2, QTableWidgetItem(sale["Customer_ID"]))
        self.sales_table.setItem(row, 3, QTableWidgetItem(sale["Product_ID"]))
        self.sales_table.setItem(row, 4, QTableWidgetItem(str(sale["Quantity"])))
        self.is_refreshing_tables = False

        self.update_status()
        self.filter_sales_table()

    def delete_selected_product(self):
        row = self.product_table.currentRow()
        if row < 0 or row >= len(self.products):
            QMessageBox.warning(self, "Error", "Select a product row to delete")
            return

        product_id = self.products[row]["Product_ID"]
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            "Delete selected product and all related sales?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.products.pop(row)
        self.sales = [s for s in self.sales if s["Product_ID"] != product_id]
        self.refresh_tables()

    def delete_selected_customer(self):
        row = self.customer_table.currentRow()
        if row < 0 or row >= len(self.customers):
            QMessageBox.warning(self, "Error", "Select a customer row to delete")
            return

        customer_id = self.customers[row]["Customer_ID"]
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            "Delete selected customer and all related sales?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.customers.pop(row)
        self.sales = [s for s in self.sales if s["Customer_ID"] != customer_id]
        self.refresh_tables()

    def delete_selected_sale(self):
        row = self.sales_table.currentRow()
        if row < 0 or row >= len(self.sales):
            QMessageBox.warning(self, "Error", "Select a sales row to delete")
            return

        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            "Delete selected sale?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.sales.pop(row)
        self.refresh_tables()

    def on_product_table_changed(self, item):
        if self.is_refreshing_tables:
            return

        row = item.row()
        col = item.column()
        if row >= len(self.products):
            return

        value = item.text().strip()

        if col == 0:
            if not value:
                QMessageBox.warning(self, "Error", "Product ID cannot be empty")
                self.refresh_tables()
                return

            for i, p in enumerate(self.products):
                if i != row and p["Product_ID"] == value:
                    QMessageBox.warning(self, "Error", "Product ID already exists")
                    self.refresh_tables()
                    return

            old_id = self.products[row]["Product_ID"]
            self.products[row]["Product_ID"] = value

            for sale in self.sales:
                if sale["Product_ID"] == old_id:
                    sale["Product_ID"] = value

        elif col == 1:
            self.products[row]["Product_Name"] = value

        elif col == 2:
            self.products[row]["Category"] = value

        elif col == 3:
            try:
                price = float(value)
                if price <= 0:
                    raise ValueError
                self.products[row]["Unit_Price"] = price
            except Exception:
                QMessageBox.warning(self, "Error", "Unit Price must be a valid positive number")
                self.refresh_tables()
                return

        elif col == 4:
            if value == "":
                self.products[row]["Cost_Price"] = None
            else:
                try:
                    cost_price = float(value)
                    if cost_price < 0:
                        raise ValueError
                    self.products[row]["Cost_Price"] = cost_price
                except Exception:
                    QMessageBox.warning(self, "Error", "Cost Price must be blank or a valid non-negative number")
                    self.refresh_tables()
                    return

        self.rebuild_dropdowns()
        self.update_status()
        self.render_charts()

    def on_customer_table_changed(self, item):
        if self.is_refreshing_tables:
            return

        row = item.row()
        col = item.column()
        if row >= len(self.customers):
            return

        value = item.text().strip()

        if col == 0:
            if not value:
                QMessageBox.warning(self, "Error", "Customer ID cannot be empty")
                self.refresh_tables()
                return

            for i, c in enumerate(self.customers):
                if i != row and c["Customer_ID"] == value:
                    QMessageBox.warning(self, "Error", "Customer ID already exists")
                    self.refresh_tables()
                    return

            old_id = self.customers[row]["Customer_ID"]
            self.customers[row]["Customer_ID"] = value

            for sale in self.sales:
                if sale["Customer_ID"] == old_id:
                    sale["Customer_ID"] = value

        elif col == 1:
            self.customers[row]["Customer_Name"] = value
        elif col == 2:
            self.customers[row]["City"] = value
        elif col == 3:
            self.customers[row]["Region"] = value

        self.rebuild_dropdowns()
        self.update_status()
        self.render_charts()

    def on_sales_table_changed(self, item):
        if self.is_refreshing_tables:
            return

        row = item.row()
        col = item.column()
        if row >= len(self.sales):
            return

        value = item.text().strip()

        if col == 0:
            if not value:
                QMessageBox.warning(self, "Error", "Order ID cannot be empty")
                self.refresh_tables()
                return

            for i, s in enumerate(self.sales):
                if i != row and s["Order_ID"] == value:
                    QMessageBox.warning(self, "Error", "Order ID already exists")
                    self.refresh_tables()
                    return

            self.sales[row]["Order_ID"] = value

        elif col == 1:
            try:
                pd.to_datetime(value)
                self.sales[row]["Date"] = value
            except Exception:
                QMessageBox.warning(self, "Error", "Invalid date format")
                self.refresh_tables()
                return

        elif col == 2:
            customer_ids = [c["Customer_ID"] for c in self.customers]
            if value not in customer_ids:
                QMessageBox.warning(self, "Error", "Customer ID does not exist")
                self.refresh_tables()
                return
            self.sales[row]["Customer_ID"] = value

        elif col == 3:
            product_ids = [p["Product_ID"] for p in self.products]
            if value not in product_ids:
                QMessageBox.warning(self, "Error", "Product ID does not exist")
                self.refresh_tables()
                return
            self.sales[row]["Product_ID"] = value

        elif col == 4:
            try:
                qty = int(value)
                if qty <= 0:
                    raise ValueError
                self.sales[row]["Quantity"] = qty
            except Exception:
                QMessageBox.warning(self, "Error", "Quantity must be a valid positive integer")
                self.refresh_tables()
                return

        self.update_status()
        self.render_charts()

    def style_axes(self, ax, title, x_label=None, y_label=None, rotate_x=0, currency_axis="y"):
        palette = self.get_chart_palette()

        ax.set_facecolor(palette["axes_bg"])
        ax.set_title(title, color=palette["text"], fontsize=13, fontweight="bold", pad=12)

        if x_label:
            ax.set_xlabel(x_label, color=palette["muted_text"], fontsize=10)
        if y_label:
            ax.set_ylabel(y_label, color=palette["muted_text"], fontsize=10)

        ax.tick_params(axis="x", colors=palette["muted_text"], labelsize=9)
        ax.tick_params(axis="y", colors=palette["muted_text"], labelsize=9)

        for label in ax.get_xticklabels():
            label.set_rotation(rotate_x)
            label.set_ha("right" if rotate_x else "center")

        ax.grid(False)

        if currency_axis == "y":
            ax.yaxis.set_major_locator(ticker.MaxNLocator(6))
            ax.yaxis.set_major_formatter(
                ticker.FuncFormatter(lambda x, pos: f"₦{int(x):,}")
            )
        elif currency_axis == "x":
            ax.xaxis.set_major_locator(ticker.MaxNLocator(6))
            ax.xaxis.set_major_formatter(
                ticker.FuncFormatter(lambda x, pos: f"₦{int(x):,}")
            )

        for spine in ax.spines.values():
            spine.set_color(palette["spine"])

    def add_bar_labels(self, ax, horizontal=False, max_val=None):
        palette = self.get_chart_palette()
        label_color = palette["text"]

        if horizontal:
            if max_val is None:
                max_val = max([patch.get_width() for patch in ax.patches], default=0)
            for patch in ax.patches:
                width = patch.get_width()
                y = patch.get_y() + patch.get_height() / 2
                ax.annotate(
                    f"₦{int(width):,}",
                    (width, y),
                    xytext=(8, 0),
                    textcoords="offset points",
                    ha="left",
                    va="center",
                    color=label_color,
                    fontsize=9,
                    clip_on=False
                )
        else:
            for patch in ax.patches:
                height = patch.get_height()
                x = patch.get_x() + patch.get_width() / 2
                ax.annotate(
                    f"₦{int(height):,}",
                    (x, height),
                    xytext=(0, 6),
                    textcoords="offset points",
                    ha="center",
                    va="bottom",
                    color=label_color,
                    fontsize=9
                )

    def add_line_labels(self, ax, x_values, y_values):
        palette = self.get_chart_palette()
        for x_val, y_val in zip(x_values, y_values):
            ax.annotate(
                f"₦{int(y_val):,}",
                (x_val, y_val),
                xytext=(0, 8),
                textcoords="offset points",
                ha="center",
                color=palette["text"],
                fontsize=9
            )

    def render_charts(self):
        palette = self.get_chart_palette()

        if not self.sales or not self.products or not self.customers:
            self.kpi_revenue.setText("₦0")
            self.kpi_orders.setText("0")
            self.kpi_units.setText("0")
            self.kpi_cost.setText("N/A")
            self.kpi_profit.setText("N/A")
            self.kpi_margin.setText("N/A")

            self.analytics_window.kpi_revenue.setText("₦0")
            self.analytics_window.kpi_orders.setText("0")
            self.analytics_window.kpi_units.setText("0")
            self.analytics_window.kpi_cost.setText("N/A")
            self.analytics_window.kpi_profit.setText("N/A")
            self.analytics_window.kpi_margin.setText("N/A")

            fig = self.analytics_window.chart_canvas.figure
            fig.clear()
            fig.patch.set_facecolor(palette["figure_bg"])
            self.analytics_window.chart_canvas.draw()
            return

        df_sales = pd.DataFrame(self.sales)
        df_products = pd.DataFrame(self.products)
        df_customers = pd.DataFrame(self.customers)

        df = df_sales.merge(df_products, on="Product_ID", how="left")
        df = df.merge(df_customers, on="Customer_ID", how="left")

        df["Revenue"] = df["Quantity"] * df["Unit_Price"]

        if "Cost_Price" not in df.columns:
            df["Cost_Price"] = pd.NA

        df["Cost"] = df["Quantity"] * df["Cost_Price"]
        df.loc[df["Cost_Price"].isna(), "Cost"] = pd.NA

        df["Profit"] = df["Revenue"] - df["Cost"]
        df.loc[df["Cost"].isna(), "Profit"] = pd.NA

        df["Margin"] = (df["Profit"] / df["Revenue"]) * 100
        df.loc[df["Cost"].isna() | (df["Revenue"] <= 0), "Margin"] = pd.NA

        df["Date"] = pd.to_datetime(df["Date"])
        df["Month"] = df["Date"].dt.month_name()
        df["Month_Num"] = df["Date"].dt.month

        revenue_by_product = df.groupby("Product_Name")["Revenue"].sum().sort_values(ascending=False)
        revenue_by_region = df.groupby("Region")["Revenue"].sum().sort_values(ascending=False)
        monthly_trend = (
            df.groupby(["Month_Num", "Month"])["Revenue"]
            .sum()
            .reset_index()
            .sort_values("Month_Num")
        )
        top_customers = (
            df.groupby("Customer_Name")["Revenue"]
            .sum()
            .sort_values(ascending=False)
            .head(5)
        )

        total_revenue = df["Revenue"].sum()
        total_orders = len(df)
        total_units = df["Quantity"].sum()

        valid_cost_rows = df[df["Cost"].notna()].copy()

        if not valid_cost_rows.empty:
            total_cost = valid_cost_rows["Cost"].sum()
            total_profit = valid_cost_rows["Profit"].sum()
            avg_margin = valid_cost_rows["Margin"].mean()
        else:
            total_cost = None
            total_profit = None
            avg_margin = None

        self.kpi_revenue.setText(f"₦{total_revenue:,.0f}")
        self.kpi_orders.setText(f"{total_orders:,}")
        self.kpi_units.setText(f"{total_units:,}")
        self.kpi_cost.setText(f"₦{total_cost:,.0f}" if total_cost is not None else "N/A")
        self.kpi_profit.setText(f"₦{total_profit:,.0f}" if total_profit is not None else "N/A")
        self.kpi_margin.setText(f"{avg_margin:.2f}%" if avg_margin is not None else "N/A")

        self.analytics_window.kpi_revenue.setText(f"₦{total_revenue:,.0f}")
        self.analytics_window.kpi_orders.setText(f"{total_orders:,}")
        self.analytics_window.kpi_units.setText(f"{total_units:,}")
        self.analytics_window.kpi_cost.setText(f"₦{total_cost:,.0f}" if total_cost is not None else "N/A")
        self.analytics_window.kpi_profit.setText(f"₦{total_profit:,.0f}" if total_profit is not None else "N/A")
        self.analytics_window.kpi_margin.setText(f"{avg_margin:.2f}%" if avg_margin is not None else "N/A")

        for lbl in [
            self.kpi_revenue, self.kpi_orders, self.kpi_units,
            self.kpi_cost, self.kpi_profit, self.kpi_margin,
            self.analytics_window.kpi_revenue,
            self.analytics_window.kpi_orders,
            self.analytics_window.kpi_units,
            self.analytics_window.kpi_cost,
            self.analytics_window.kpi_profit,
            self.analytics_window.kpi_margin
        ]:
            self.animate_kpi(lbl)

        fig = self.analytics_window.chart_canvas.figure
        fig.clear()
        fig.patch.set_facecolor(palette["figure_bg"])

        ax1 = fig.add_subplot(221)
        ax2 = fig.add_subplot(222)
        ax3 = fig.add_subplot(223)
        ax4 = fig.add_subplot(224)

        revenue_by_product.plot(
            kind="bar",
            ax=ax1,
            color=palette["product"],
            edgecolor=palette["product_edge"],
            linewidth=0.6
        )
        self.style_axes(ax1, "Revenue by Product", y_label="Revenue", rotate_x=20, currency_axis="y")
        self.add_bar_labels(ax1, horizontal=False)

        revenue_by_region.plot(
            kind="bar",
            ax=ax2,
            color=palette["region"],
            edgecolor=palette["region_edge"],
            linewidth=0.6
        )
        self.style_axes(ax2, "Revenue by Region", y_label="Revenue", rotate_x=12, currency_axis="y")
        self.add_bar_labels(ax2, horizontal=False)

        ax3.plot(
            monthly_trend["Month"],
            monthly_trend["Revenue"],
            marker="o",
            color=palette["trend"],
            linewidth=2.4,
            markersize=7,
            markerfacecolor=palette["trend_marker"]
        )
        self.style_axes(ax3, "Monthly Revenue Trend", y_label="Revenue", rotate_x=0, currency_axis="y")
        self.add_line_labels(ax3, monthly_trend["Month"], monthly_trend["Revenue"])

        top_customers.plot(
            kind="barh",
            ax=ax4,
            color=palette["customer"],
            edgecolor=palette["customer_edge"],
            linewidth=0.6
        )

        max_val = top_customers.max()
        ax4.set_xlim(0, max_val * 1.25)
        ax4.margins(y=0.20)

        self.style_axes(ax4, "Top Customers", x_label="Revenue", currency_axis="x")
        self.add_bar_labels(ax4, horizontal=True, max_val=max_val)

        fig.subplots_adjust(hspace=0.42, wspace=0.28, left=0.08, right=0.97, top=0.93, bottom=0.10)
        self.analytics_window.chart_canvas.draw()

    def open_dashboard(self):
        self.analytics_window.show()

    def export_dashboard_png(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Dashboard as PNG",
            "dashboard.png",
            "PNG Files (*.png)"
        )

        if not file_path:
            return

        self.analytics_window.chart_canvas.figure.savefig(
            file_path,
            dpi=200,
            facecolor=self.analytics_window.chart_canvas.figure.get_facecolor(),
            bbox_inches="tight"
        )

        QMessageBox.information(self, "Export Successful", f"Dashboard PNG saved to:\n{file_path}")

    def export_excel_report(self):
        company = self.company_input.text().strip()

        if not company:
            QMessageBox.warning(self, "Error", "Enter company name")
            return

        if not self.products or not self.customers or not self.sales:
            QMessageBox.warning(self, "Error", "Incomplete data")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Excel Report",
            f"{company}_Business_Analytics_Report.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not file_path:
            return

        self.status_output.setText("Exporting Excel report...")
        run_full_analysis(self.products, self.customers, self.sales, company, file_path)
        self.status_output.setText("Excel report exported successfully.")
        QMessageBox.information(self, "Export Successful", f"Excel report saved to:\n{file_path}")

    def generate_report(self):
        company = self.company_input.text().strip()

        if not company:
            QMessageBox.warning(self, "Error", "Enter company name")
            return

        if not self.products or not self.customers or not self.sales:
            QMessageBox.warning(self, "Error", "Incomplete data")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Excel Report",
            f"{company}_Business_Analytics_Report.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not file_path:
            return

        self.status_output.setText("Running enterprise analytics...")
        run_full_analysis(self.products, self.customers, self.sales, company, file_path)
        self.render_charts()
        self.analytics_window.show()

        self.status_output.setText("Analytics completed. Excel dashboard generated.")
        QMessageBox.information(
            self,
            "Success",
            f"Dashboard generated successfully.\nExcel saved to:\n{file_path}"
        )

    def save_project_data(self):
        company = self.company_input.text().strip()

        if not company:
            QMessageBox.warning(self, "Error", "Enter company name first")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Project File",
            f"{company}_project.json",
            "JSON Files (*.json)"
        )

        if not file_path:
            return

        save_project(file_path, company, self.products, self.customers, self.sales)
        QMessageBox.information(self, "Saved", f"Project saved successfully to:\n{file_path}")

    def load_project_data(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Load Project File",
            "",
            "JSON Files (*.json)"
        )

        if not file_path:
            return

        data = load_project(file_path)

        if not data:
            QMessageBox.warning(self, "Error", "Could not load selected project file")
            return

        self.company_input.setText(data["company"])
        self.products = data["products"]
        self.customers = data["customers"]
        self.sales = data["sales"]

        self.refresh_tables()
        QMessageBox.information(self, "Loaded", f"Project loaded successfully from:\n{file_path}")

    def refresh_tables(self):
        self.is_refreshing_tables = True

        self.product_table.setRowCount(0)
        self.customer_table.setRowCount(0)
        self.sales_table.setRowCount(0)

        self.rebuild_dropdowns()

        for p in self.products:
            row = self.product_table.rowCount()
            self.product_table.insertRow(row)
            self.product_table.setItem(row, 0, QTableWidgetItem(p["Product_ID"]))
            self.product_table.setItem(row, 1, QTableWidgetItem(p["Product_Name"]))
            self.product_table.setItem(row, 2, QTableWidgetItem(p["Category"]))
            self.product_table.setItem(row, 3, QTableWidgetItem(str(p["Unit_Price"])))
            self.product_table.setItem(
                row, 4,
                QTableWidgetItem("" if p.get("Cost_Price") is None else str(p.get("Cost_Price")))
            )

        for c in self.customers:
            row = self.customer_table.rowCount()
            self.customer_table.insertRow(row)
            self.customer_table.setItem(row, 0, QTableWidgetItem(c["Customer_ID"]))
            self.customer_table.setItem(row, 1, QTableWidgetItem(c["Customer_Name"]))
            self.customer_table.setItem(row, 2, QTableWidgetItem(c["City"]))
            self.customer_table.setItem(row, 3, QTableWidgetItem(c["Region"]))

        for s in self.sales:
            row = self.sales_table.rowCount()
            self.sales_table.insertRow(row)
            self.sales_table.setItem(row, 0, QTableWidgetItem(s["Order_ID"]))
            self.sales_table.setItem(row, 1, QTableWidgetItem(s["Date"]))
            self.sales_table.setItem(row, 2, QTableWidgetItem(s["Customer_ID"]))
            self.sales_table.setItem(row, 3, QTableWidgetItem(s["Product_ID"]))
            self.sales_table.setItem(row, 4, QTableWidgetItem(str(s["Quantity"])))

        self.is_refreshing_tables = False
        self.update_status()
        self.render_charts()

        self.filter_product_table()
        self.filter_customer_table()
        self.filter_sales_table()


def run_gui():
    app = QApplication(sys.argv)
    window = BusinessAnalyticsApp()
    window.show()
    sys.exit(app.exec_())