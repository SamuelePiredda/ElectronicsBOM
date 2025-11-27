import sys
import os
import time
import requests
import cloudscraper 
import webbrowser 
import re 
import csv 
from datetime import datetime 
from bs4 import BeautifulSoup

# --- XHTML2PDF FOR PDF GENERATION ---
# Imports the library required to convert the dynamically generated HTML into a PDF file.
try:
    from xhtml2pdf import pisa
except ImportError:
    print("ERROR: You must install xhtml2pdf! Run: pip install xhtml2pdf")
    sys.exit(1)

# SQLAlchemy imports for ORM (Object Relational Mapper) database management
from sqlalchemy import create_engine, Column, Integer, String, Text, ForeignKey, Float
from sqlalchemy.orm import declarative_base, sessionmaker, relationship

# PyQt6 imports for the Graphical User Interface (GUI)
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QTableWidget, QTableWidgetItem, QPushButton, QLabel, QLineEdit, 
                             QSplitter, QListWidget, QListWidgetItem, QDialog, QFormLayout, 
                             QMessageBox, QTextEdit, QAbstractItemView, QHeaderView, QInputDialog,
                             QComboBox, QMenu, QFileDialog)
from PyQt6.QtCore import Qt, QRunnable, QThreadPool, pyqtSignal, QObject, pyqtSlot, QSettings
from PyQt6.QtGui import QColor, QPalette, QAction, QFont, QIcon
import qtawesome as qta  # Library for vector icons (FontAwesome)

# =============================================================================
# 1. CURRENCY MANAGER
# =============================================================================
class CurrencyManager:
    """
    Manages the USD -> EUR exchange rate.
    Uses a caching mechanism to prevent excessive API calls.
    """
    _rate = None
    _last_update = 0
    
    @staticmethod
    def get_usd_to_eur():
        """
        Retrieves the current exchange rate.
        If the cached rate is older than 24 hours (86400s), it fetches a new one.
        """
        now = time.time()
        # Update if rate is missing or older than 1 day
        if CurrencyManager._rate is None or (now - CurrencyManager._last_update) > 86400:
            try:
                # Free public API for exchange rates
                url = "https://open.er-api.com/v6/latest/USD"
                response = requests.get(url, timeout=5)
                if response.status_code == 200:
                    data = response.json()
                    CurrencyManager._rate = data['rates']['EUR']
                    CurrencyManager._last_update = now
            except Exception as e:
                # Fallback in case of network error
                if CurrencyManager._rate is None: 
                    CurrencyManager._rate = 0.92  # Default static fallback
                    QMessageBox.warning(None, "Error retrieving €/$ exchange rate", f"Error: {e}")
                    
        return CurrencyManager._rate

# =============================================================================
# 2. DATABASE MODELS (SQLAlchemy)
# =============================================================================
Base = declarative_base()

class Project(Base):
    """
    Table to store Projects.
    """
    __tablename__ = 'projects'
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True) # Project name must be unique
    notes = Column(Text, default="")   # Free text notes
    # One-to-Many relationship: One project has many components
    components = relationship("Component", back_populates="project", cascade="all, delete-orphan")

class Component(Base):
    """
    Table to store individual Electronic Components.
    """
    __tablename__ = 'components'
    id = Column(Integer, primary_key=True)
    project_id = Column(Integer, ForeignKey('projects.id')) # Foreign Key to Project
    
    mouser_part_number = Column(String) # Code for Mouser API
    jlc_part_number = Column(String)    # Code for LCSC/JLCPCB scraping
    description = Column(String)
    category = Column(String, default="Other")
    target_qty = Column(Integer, default=1) # Required quantity per board
    backup_part = Column(String, nullable=True)
    
    # Cache Fields: Store the last fetched data for offline/fast usage
    last_mouser_stock = Column(Integer, default=-1)
    last_mouser_price = Column(Float, default=0.0) 
    last_jlc_stock = Column(Integer, default=-1)
    last_jlc_price = Column(Float, default=0.0)    
    last_update = Column(String, default="") 
    
    project = relationship("Project", back_populates="components")

# Categories for the dropdown menu
CATEGORIES = ["All", "Resistor", "Capacitor", "Inductor", "IC", "Microcontroller", 
              "Connector", "Transistor", "Diode", "Sensor", "Module", "Other"] 

def init_db(db_path):
    """
    Initializes the connection to the SQLite database.
    Creates tables if they do not exist.
    Returns a session factory to interact with the DB.
    """
    if not os.path.exists(db_path):
        pass # The file will be created by create_engine if it doesn't exist

    db_url = f"sqlite:///{db_path}"
    engine = create_engine(db_url)
    Base.metadata.create_all(engine) # Generates schema
    return sessionmaker(bind=engine)

# =============================================================================
# 3. SEARCH FUNCTIONS (SCRAPING & API)
# =============================================================================

def get_jlcpcb_stats(code, qnty):
    """
    Scrapes the JLCPCB part detail page to get stock and pricing.
    Input: LCSC Code (e.g., C1234), Required Total Quantity.
    Output: [Available Stock, Total Price in EUR]
    """
    if not code: return [-1, 0.0]
    headers = {'User-Agent': 'Mozilla/5.0 ...'} # User agent to avoid basic blocks
    url = "https://jlcpcb.com/partdetail/" + code
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status() 
    except: return [-1, 0.0]

    soup = BeautifulSoup(response.text, 'html.parser')
    
    # 1. Extract Stock
    stock_element = soup.find('div', class_='text-16 font-bold') 
    quantity = 0; total_price_usd = 0.0

    if stock_element:
        try:
            stock_text = stock_element.get_text(strip=True)
            # Cleans string to find the number (e.g., "Stock: 12,000")
            if ":" in stock_text: quantity = int(stock_text.replace(',', '').split(":")[1])
            else: quantity = int(''.join(filter(str.isdigit, stock_text)))
        except: quantity = 0
    
    # 2. Extract Price (Tiered Pricing Table)
    # Finds the div containing prices
    cost_div = soup.find('div', class_='mt-10 bg-[#f8f8f8] py-10 px-30')
    if cost_div:
        cost_text = cost_div.get_text()        
        # Regex to find quantities (e.g., 1+, 10+, 100+)
        quant = list(map(lambda x: int(x.replace(',', '')), re.findall(r'(\d{1,3}(?:\d{3})*)\+', cost_text)))
        # Regex to find prices (e.g., $0.50)
        cost = list(map(float, re.findall(r'\$(\d+\.\d+)', cost_text)))
        
        # Maps quantity to price and sorts
        tiers = sorted(zip(quant, cost), key=lambda x: x[0])
        if tiers:
            target_unit_price = tiers[0][1] # Base price
            # Finds the correct tier for the requested quantity
            for qty_tier, price_tier in tiers:
                if qnty >= qty_tier: target_unit_price = price_tier
                else: break 
            total_price_usd = qnty * target_unit_price

    # Converts to Euro before returning
    return [quantity, total_price_usd * CurrencyManager.get_usd_to_eur()]

def get_mouser_stats(part_number, qty, api_key):
    """
    Queries the Mouser Search API.
    Input: Part Number, Quantity, API Key.
    Output: [Stock, Total Price in EUR]
    """
    if not part_number or not api_key: return [-1, 0.0]
    url = f"https://api.mouser.com/api/v1/search/keyword?apiKey={api_key}"
    headers = {'Content-Type': 'application/json'}
    # Payload for Mouser Search API v1
    body = {"SearchByKeywordRequest": {"keyword": part_number, "records": 1, "startingRecord": 0, "searchOptions": "None"}}
    
    try:
        r = requests.post(url, json=body, headers=headers, timeout=5)
        if r.status_code == 200:
            data = r.json()
            if data.get('Errors'): return [-1, 0.0]
            
            results = data.get('SearchResults', {}).get('Parts', [])
            if results:
                part = results[0] # Takes the first result
                
                # Parsing Stock
                avail_str = str(part.get('Availability', '0'))
                if avail_str == 'None': avail_str = str(part.get('FactoryStock', '0'))
                stock = int(''.join(filter(str.isdigit, avail_str)) or 0)
                
                # Parsing Price (PriceBreaks)
                price_unit = 0.0
                breaks = part.get('PriceBreaks', [])
                for pb in breaks:
                    # Finds the correct price tier
                    if qty >= int(pb.get('Quantity', 99999)):
                        try: price_unit = float(pb.get('Price', '0').replace('€','').replace('$','').replace(',','.').strip())
                        except: pass
                
                # If no tier found, use the first available one
                if price_unit == 0 and breaks:
                     try: price_unit = float(breaks[0].get('Price', '0').replace('€','').replace('$','').replace(',','.').strip())
                     except: pass
                     
                return [stock, price_unit * qty]
    except: pass
    return [-1, 0.0]

# =============================================================================
# 4. UPDATE THREAD (Multithreading)
# =============================================================================
class WorkerSignals(QObject):
    """Signals to communicate from the worker thread to the GUI."""
    result = pyqtSignal(dict) 

class DataUpdater(QRunnable):
    """
    Worker executed in a separate thread (QThreadPool).
    Prevents GUI freezing during network requests.
    """
    def __init__(self, component_id, mouser_pn, jlc_pn, qty):
        super().__init__()
        self.c_id = component_id; self.m_pn = mouser_pn; self.j_pn = jlc_pn; self.qty = qty
        self.signals = WorkerSignals(); self.settings = QSettings("MySoft", "BOMManager")

    @pyqtSlot()
    def run(self):
        """Executes requests to Mouser and JLCPCB conceptually in parallel."""
        m_res = get_mouser_stats(self.m_pn, self.qty, self.settings.value("mouser_key", "").strip())
        j_res = get_jlcpcb_stats(self.j_pn, self.qty)
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        # Emits data to GUI via signal
        self.signals.result.emit({
            'id': self.c_id, 
            'mouser_stock': m_res[0], 'mouser_price': m_res[1], 
            'jlc_stock': j_res[0], 'jlc_price': j_res[1],
            'timestamp': now_str
        })

# =============================================================================
# 5. UI: NOTES DIALOG
# =============================================================================
class NotesDialog(QDialog):
    """Simple window to edit project text notes."""
    def __init__(self, parent=None, project=None):
        super().__init__(parent)
        self.setWindowTitle(f"Project Notes: {project.name}")
        self.resize(500, 400)
        self.project = project
        self.session = parent.session
        lay = QVBoxLayout(self)
        self.txt = QTextEdit(); self.txt.setText(project.notes)
        lay.addWidget(self.txt)
        btn = QPushButton("Save and Close"); btn.clicked.connect(self.save_and_close)
        lay.addWidget(btn)

    def save_and_close(self):
        self.project.notes = self.txt.toPlainText()
        self.session.commit() # Save to DB
        self.accept()

# =============================================================================
# 6. UI: COMPONENTS AND SETTINGS
# =============================================================================
class ComponentDialog(QDialog):
    """
    Dialog to Add or Edit a component.
    Shows fields for PN, Description, Quantity, etc.
    """
    def __init__(self, parent=None, component=None):
        super().__init__(parent)
        self.setWindowTitle("Component Details")
        self.resize(500, 450)
        self.layout = QFormLayout(self)
        
        # Widget Definitions
        self.inp_cat = QComboBox(); self.inp_cat.addItems(CATEGORIES[1:])
        self.inp_m_pn = QLineEdit()
        self.inp_j_pn = QLineEdit(); self.inp_j_pn.setPlaceholderText("Ex. C7593")
        self.inp_desc = QLineEdit()
        self.inp_qty = QLineEdit("1")
        self.inp_backup = QLineEdit()
        
        # If editing, populate fields with existing data
        if component:
            self.inp_m_pn.setText(component.mouser_part_number.replace("\n","") if component.mouser_part_number else "")
            self.inp_j_pn.setText(component.jlc_part_number.replace("\n","") if component.jlc_part_number else "")
            self.inp_desc.setText(component.description.replace("\n","") if component.description else "")
            self.inp_qty.setText(str(component.target_qty))
            self.inp_backup.setText(component.backup_part or "")
            idx = self.inp_cat.findText(component.category)
            if idx >= 0: self.inp_cat.setCurrentIndex(idx)
        
        # Add rows to form
        self.layout.addRow("Category:", self.inp_cat)
        self.layout.addRow("Mouser PN:", self.inp_m_pn)
        self.layout.addRow("JLCPCB Code:", self.inp_j_pn)
        self.layout.addRow("Description:", self.inp_desc)
        self.layout.addRow("Quantity:", self.inp_qty)
        self.layout.addRow("Backup:", self.inp_backup)
        
        # Quick link buttons to check component on browser
        ll = QHBoxLayout()
        b_m = QPushButton(qta.icon('fa5s.external-link-alt'), "Open Mouser"); b_m.clicked.connect(lambda: self.open_l(f"https://www.mouser.it/c/?q={self.inp_m_pn.text()}"))
        b_j = QPushButton(qta.icon('fa5s.external-link-alt'), "Open JLCPCB"); b_j.clicked.connect(lambda: self.open_l(f"https://jlcpcb.com/partdetail/{self.inp_j_pn.text()}"))
        ll.addWidget(b_m); ll.addWidget(b_j)
        self.layout.addRow(QLabel("Links:"), ll)

        # Action Buttons (Save/Cancel)
        bb = QHBoxLayout()
        bs = QPushButton("Save"); bs.clicked.connect(self.accept)
        bc = QPushButton("Cancel"); bc.clicked.connect(self.reject)
        bb.addStretch(); bb.addWidget(bc); bb.addWidget(bs)
        self.layout.addRow(bb)

    def open_l(self, url): 
        # Opens browser only if URL looks valid
        if url.split('=')[-1] and url.split('/')[-1]: webbrowser.open(url)

    def get_data(self):
        """Returns a dictionary with user inputs."""
        return {'cat': self.inp_cat.currentText(), 'm_pn': self.inp_m_pn.text().replace("\n",""), 'j_pn': self.inp_j_pn.text().replace("\n",""), 'desc': self.inp_desc.text().replace("\n",""), 'qty': int(self.inp_qty.text()) if self.inp_qty.text().isdigit() else 1, 'backup': self.inp_backup.text()}

class SettingsDialog(QDialog):
    """
    Dialog to manage global settings:
    1. Mouser API Key
    2. Database Path (create or switch file)
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.resize(500, 250)
        self.settings = QSettings("MySoft", "BOMManager")
        
        layout = QVBoxLayout(self)
        form = QFormLayout()
        
        # Mouser API Key Field
        self.i_m = QLineEdit(self.settings.value("mouser_key", ""))
        form.addRow("Mouser API Key:", self.i_m)
        
        # Current DB Path Display
        current_db = self.settings.value("db_path", "Not Set")
        self.lbl_db = QLabel(current_db)
        self.lbl_db.setStyleSheet("color: gray; font-size: 10px; font-family: monospace;")
        self.lbl_db.setWordWrap(True)
        form.addRow("Current Database:", self.lbl_db)
        
        layout.addLayout(form)
        
        # DB Management Buttons
        db_btn_layout = QHBoxLayout()
        btn_open_db = QPushButton("Select Existing DB")
        btn_open_db.clicked.connect(self.select_existing_db)
        btn_open_db.setToolTip("Use this if the DB file already exists (e.g. on Dropbox/Drive)")
        
        btn_new_db = QPushButton("Create New DB")
        btn_new_db.clicked.connect(self.create_new_db)
        
        db_btn_layout.addWidget(btn_open_db)
        db_btn_layout.addWidget(btn_new_db)
        layout.addLayout(db_btn_layout)
        
        layout.addStretch()
        
        # Save Button
        b_save = QPushButton("Save and Close")
        b_save.clicked.connect(self.save_settings)
        layout.addWidget(b_save)

    def select_existing_db(self):
        """Allows selecting an existing .db file."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Existing Database", "", "SQLite Database (*.db)")
        if file_path:
            self.settings.setValue("db_path", file_path)
            self.lbl_db.setText(file_path)
            QMessageBox.warning(self, "Restart Required", "Path updated.\nPlease restart the application to load the new database.")

    def create_new_db(self):
        """Allows creating a new empty .db file."""
        file_path, _ = QFileDialog.getSaveFileName(self, "Create New Database", "my_bom_manager_new.db", "SQLite Database (*.db)")
        if file_path:
            self.settings.setValue("db_path", file_path)
            self.lbl_db.setText(file_path)
            QMessageBox.warning(self, "Restart Required", "New path set.\nPlease restart the application to create and use the new database.")

    def save_settings(self):
        self.settings.setValue("mouser_key", self.i_m.text().strip())
        self.accept()

# =============================================================================
# 7. MAIN WINDOW
# =============================================================================
class MainWindow(QMainWindow):
    def __init__(self, session_factory):
        super().__init__()
        self.setWindowTitle("BOM Manager Pro")
        self.setWindowIcon(QIcon("icon.ico")) # App Icon
        self.resize(1350, 850)
        self.session = session_factory() # Create DB session
        self.threadpool = QThreadPool()  # Thread pool for update workers
        self.current_project = None
        
        # Totals Variables
        self.tm = 0.0 # Mouser Total
        self.tj = 0.0 # JLCPCB Total
        self.hybrid_total = 0.0 # "Best Price" Total
        
        self.init_ui()
        self.load_projects()

    def init_ui(self):
        """Configures the graphical layout: Horizontal Splitter (Project List | BOM Table)."""
        mw = QWidget(); self.setCentralWidget(mw)
        ml = QHBoxLayout(mw); ml.setContentsMargins(0,0,0,0)
        sp = QSplitter(Qt.Orientation.Horizontal); ml.addWidget(sp)

        # --- LEFT PANEL (Project List) ---
        left = QWidget(); ll = QVBoxLayout(left); ll.setContentsMargins(0,0,0,0)
        ll.addWidget(QLabel("  PROJECTS"))
        self.p_list = QListWidget()
        self.p_list.itemClicked.connect(self.select_project)
        self.p_list.setFont(QFont("sans-serif", 10))
        ll.addWidget(self.p_list)
        
        # New/Delete Project Buttons
        pb = QHBoxLayout()
        ba = QPushButton("New"); ba.clicked.connect(self.add_project)
        bd = QPushButton("Delete"); bd.clicked.connect(self.delete_project)
        pb.addWidget(ba); pb.addWidget(bd)
        ll.addLayout(pb)
        sp.addWidget(left)

        # --- RIGHT PANEL (BOM and Details) ---
        right = QWidget(); rl = QVBoxLayout(right); rl.setContentsMargins(0,0,0,0)
        
        # Right Header (Project Title, Toolbar)
        head = QWidget(); hl = QHBoxLayout(head); head.setStyleSheet("background:#222; padding:5px;")
        
        self.lbl_t = QLabel("Select Project"); self.lbl_t.setStyleSheet("font-size:18px; font-weight:bold; border:none; background:none;")
        hl.addWidget(self.lbl_t)
        
        # Toolbar Buttons (Edit Name, Notes, Export, Filters, Settings, Refresh)
        btn_edit = QPushButton(qta.icon('fa5s.pen'), "")
        btn_edit.setToolTip("Rename Project"); btn_edit.setFixedSize(30, 30); btn_edit.clicked.connect(self.rename_project)
        hl.addWidget(btn_edit)

        btn_notes = QPushButton(qta.icon('fa5s.sticky-note', color="#FFC107"), "")
        btn_notes.setToolTip("Project Notes"); btn_notes.setFixedSize(30, 30); btn_notes.clicked.connect(self.open_notes)
        hl.addWidget(btn_notes)
        hl.addStretch()

        # Export Menu (CSV/PDF)
        btn_exp = QPushButton(qta.icon('fa5s.file-export'), "Export")
        menu_exp = QMenu()
        act_csv = QAction("Export CSV (Detailed)", self); act_csv.triggered.connect(self.export_csv)
        act_pdf = QAction("Export PDF (Landscape)", self); act_pdf.triggered.connect(self.export_pdf)
        menu_exp.addAction(act_csv); menu_exp.addAction(act_pdf)
        btn_exp.setMenu(menu_exp)
        hl.addWidget(btn_exp)
        
        # Category Filter
        self.cmb_f = QComboBox(); self.cmb_f.addItems(CATEGORIES)
        self.cmb_f.currentTextChanged.connect(self.apply_filter)
        hl.addWidget(QLabel("Filter:")); hl.addWidget(self.cmb_f); hl.addSpacing(10)
        
        bs = QPushButton(qta.icon('fa5s.cog'), ""); bs.clicked.connect(lambda: SettingsDialog(self).exec())
        hl.addWidget(bs)
        br = QPushButton(qta.icon('fa5s.sync'), "Refresh"); br.clicked.connect(self.refresh_prices)
        hl.addWidget(br)
        rl.addWidget(head)

        # TABLE SETUP (QTableWidget)
        self.tab = QTableWidget()
        self.cols = ["ID", "Mouser PN", "JLC Code", "Cat", "Desc", "Qty", "Mouser", "JLCPCB"]
        self.tab.setColumnCount(len(self.cols)); self.tab.setHorizontalHeaderLabels(self.cols)
        self.tab.verticalHeader().setVisible(False)
        self.tab.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tab.doubleClicked.connect(self.edit_component)
        self.tab.setStyleSheet("QTableWidget{background:#1e1e1e; color:#ddd; gridline-color:#333;}")
        
        # Text display optimizations
        self.tab.setWordWrap(True) # Wrap text
        self.tab.setTextElideMode(Qt.TextElideMode.ElideNone) # Disable "..." elision
        
        header = self.tab.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents) # Fit to content
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch) # Description takes remaining space

        rl.addWidget(self.tab)

        # Component Management Buttons (Add/Remove)
        act = QHBoxLayout()
        bac = QPushButton("Add Component"); bac.clicked.connect(self.add_component)
        brc = QPushButton("Remove"); brc.clicked.connect(self.del_component)
        act.addWidget(bac); act.addWidget(brc); act.addStretch()
        rl.addLayout(act)

        # Footer with Totals
        self.lbl_stat = QLabel("Totals: -"); self.lbl_stat.setStyleSheet("background:#007acc; padding:10px; font-weight:bold;")
        rl.addWidget(self.lbl_stat)
        
        # Set splitter proportions
        sp.addWidget(right); sp.setSizes([200, 1100])

    # --- HELPERS ---
    def get_last_refresh_date(self):
        """Finds the most recent update date among all components."""
        dates = []
        if not self.current_project: return "Never"
        for c in self.current_project.components:
            if c.last_update: dates.append(c.last_update)
        return max(dates) if dates else "Never"

    def calculate_unit(self, total, qty):
        """Calculates unit price."""
        if qty <= 0: return 0.0
        return total / qty

    # --- PROJECTS & NOTES LOGIC ---
    def rename_project(self):
        """Renames the current project in the DB."""
        if not self.current_project: return
        new_n, ok = QInputDialog.getText(self, "Rename", "New name:", text=self.current_project.name)
        if ok and new_n:
            try:
                self.current_project.name = new_n; self.session.commit()
                self.lbl_t.setText(new_n); self.load_projects()
            except Exception as e: QMessageBox.warning(self, "Error", f"Error (duplicate name?): {e}")

    def open_notes(self):
        """Opens the notes dialog."""
        if not self.current_project: return
        NotesDialog(self, self.current_project).exec()

    # --- EXPORT CSV ---
    def export_csv(self):
        """Exports the current BOM to a CSV file."""
        if not self.current_project: return
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", f"{self.current_project.name}.csv", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(["PROJECT", self.current_project.name])
                    writer.writerow(["GENERATED ON", datetime.now().strftime("%Y-%m-%d %H:%M")])
                    writer.writerow(["LAST REFRESH", self.get_last_refresh_date()])
                    writer.writerow([]) 
                    
                    headers = ["Mouser PN", "JLC Code", "Category", "Description", "Qty", 
                               "Mouser Stock", "Mouser Unit (€)", "Mouser Total (€)",
                               "JLC Stock", "JLC Unit (€)", "JLC Total (€)"]
                    writer.writerow(headers)
                    
                    for c in self.current_project.components:
                        # Logic to handle N/A values
                        m_ok = c.last_mouser_stock >= c.target_qty
                        m_stk = c.last_mouser_stock if c.last_mouser_stock > -1 else "N/A"
                        m_tot = c.last_mouser_price if m_ok else 0.0
                        m_unit = self.calculate_unit(m_tot, c.target_qty) if m_ok else 0.0
                        
                        j_ok = c.last_jlc_stock >= c.target_qty
                        j_stk = c.last_jlc_stock if c.last_jlc_stock > -1 else "N/A"
                        j_tot = c.last_jlc_price if j_ok else 0.0
                        j_unit = self.calculate_unit(j_tot, c.target_qty) if j_ok else 0.0
                        
                        writer.writerow([
                            c.mouser_part_number, c.jlc_part_number, c.category, c.description, c.target_qty,
                            m_stk, f"{m_unit:.3f}", f"{m_tot:.2f}",
                            j_stk, f"{j_unit:.3f}", f"{j_tot:.2f}"
                        ])
                QMessageBox.information(self, "Export", "CSV Saved!")
            except Exception as e: QMessageBox.critical(self, "Error", str(e))

    # --- EXPORT PDF (COMPACT CARD LAYOUT) ---
    def export_pdf(self):
        """
        Generates a PDF formatted with a "Card Slim" layout.
        Constructs HTML/CSS strings and uses xhtml2pdf for conversion.
        """
        if not self.current_project: return
        path, _ = QFileDialog.getSaveFileName(self, "Save PDF", f"{self.current_project.name}.pdf", "PDF Files (*.pdf)")
        if not path: return

        cards_html = ""
        
        for c in self.current_project.components:
            # Availability calculations
            m_ok = c.last_mouser_stock >= c.target_qty
            j_ok = c.last_jlc_stock >= c.target_qty
            
            # Mouser Data Formatting (Green if OK, Red if No Stock)
            m_stk_style = "color: #2e7d32; font-weight:bold;" if m_ok else "color: #c62828;"
            m_stk_txt = f"{c.last_mouser_stock}" if c.last_mouser_stock > -1 else "N/A"
            m_price = f"{c.last_mouser_price:.2f} €" if m_ok else "-"
            m_unit = f"{self.calculate_unit(c.last_mouser_price, c.target_qty):.3f}" if m_ok else "-"
            
            # JLC Data Formatting
            j_stk_style = "color: #1565c0; font-weight:bold;" if j_ok else "color: #c62828;"
            j_stk_txt = f"{c.last_jlc_stock}" if c.last_jlc_stock > -1 else "N/A"
            j_price = f"{c.last_jlc_price:.2f} €" if j_ok else "-"
            j_unit = f"{self.calculate_unit(c.last_jlc_price, c.target_qty):.3f}" if j_ok else "-"

            # "Winner" logic (lowest price with stock) to highlight the cell
            winner_m = False
            winner_j = False
            if m_ok and j_ok:
                if c.last_mouser_price < c.last_jlc_price: winner_m = True
                else: winner_j = True
            elif m_ok: winner_m = True
            elif j_ok: winner_j = True
            
            m_bg = "#e8f5e9" if not winner_m else "#c8e6c9" 
            j_bg = "#e3f2fd" if not winner_j else "#bbdefb" 
            m_border = "2px solid #4caf50" if winner_m else "1px solid #ddd"
            j_border = "2px solid #2196f3" if winner_j else "1px solid #ddd"

            # HTML construction of the component "Card"
            cards_html += f"""
            <div class="card">
                <div class="card-header">
                    <table width="100%">
                        <tr>
                            <td width="85%">
                                <div><span class="cat-badge">{c.category}</span> <span class="desc">{c.description}</span></div>
                            </td>
                            <td width="15%" align="right">
                                <div class="qty-line"><b>x{c.target_qty}</b></div>
                            </td>
                        </tr>
                    </table>
                </div>

                <table class="card-body" cellspacing="2">
                    <tr>
                        <td width="50%" class="vendor-box" style="background-color: {m_bg}; border: {m_border};">
                            <div class="row-flex">
                                <span class="v-title" style="color:#2e7d32;">MOUSER</span>
                                <span class="pn">{c.mouser_part_number}</span>
                            </div>
                            <div class="metrics-row">
                                Stk: <span style="{m_stk_style}">{m_stk_txt}</span> | Unit: {m_unit}
                            </div>
                            <div class="total-price">TOT: {m_price}</div>
                        </td>

                        <td width="50%" class="vendor-box" style="background-color: {j_bg}; border: {j_border};">
                            <div class="row-flex">
                                <span class="v-title" style="color:#1565c0;">JLCPCB</span>
                                <span class="pn">{c.jlc_part_number}</span>
                            </div>
                            <div class="metrics-row">
                                Stk: <span style="{j_stk_style}">{j_stk_txt}</span> | Unit: {j_unit}
                            </div>
                            <div class="total-price">TOT: {j_price}</div>
                        </td>
                    </tr>
                </table>
            </div>
            """

        notes_html = ""
        if self.current_project.notes:
            notes_html = f"<div class='notes'><h3>Notes:</h3><p>{self.current_project.notes.replace(chr(10), '<br>')}</p></div>"

        # Complete HTML structure with embedded CSS
        html_content = f"""
        <html>
        <head>
            <style>
                @page {{ size: a4 portrait; margin: 1cm; }}
                body {{ font-family: Helvetica, Arial, sans-serif; color: #333; font-size: 9px; }}
                /* CSS Omitted for brevity (same as logic above) */
                h1 {{ font-size: 18px; color: #222; border-bottom: 2px solid #007acc; padding-bottom: 5px; margin-bottom: 10px; }}
                .meta {{ font-size: 9px; color: #666; margin-bottom: 15px; }}
                .card {{ border: 1px solid #ccc; background-color: #fff; margin-bottom: 8px; page-break-inside: avoid; }}
                .card-header {{ background-color: #f0f0f0; padding: 3px 5px; border-bottom: 1px solid #ddd; }}
                .desc {{ font-size: 10px; font-weight: bold; color: #000; word-wrap: break-word; }}
                .cat-badge {{ font-size: 8px; background: #666; color: white; padding: 1px 4px; border-radius: 3px; font-weight: normal; margin-right: 5px; }}
                .qty-line {{ font-size: 11px; color: #000; text-align: right; }}
                .card-body {{ width: 100%; }}
                .vendor-box {{ padding: 3px; vertical-align: top; }}
                .row-flex {{ margin-bottom: 2px; border-bottom: 1px solid rgba(0,0,0,0.05); padding-bottom: 1px; }}
                .v-title {{ font-size: 9px; font-weight: bold; margin-right: 5px; }}
                .pn {{ font-family: Consolas; font-size: 9px; color: #444; word-wrap: break-word; font-weight: bold;}}
                .metrics-row {{ font-size: 9px; color: #333; margin-bottom: 1px; }}
                .total-price {{ font-size: 10px; font-weight: bold; text-align: right; margin-top: 1px; }}
                .notes {{ background: #ffffea; border: 1px solid #e0e0a0; padding: 5px; font-size: 9px; }}
            </style>
        </head>
        <body>
            <h1>BOM: {self.current_project.name}</h1>
            <div class="meta">
                Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Data updated: {self.get_last_refresh_date()}
            </div>
            
            {cards_html}
            
            <hr>
            <h3><b>MOUSER: {self.tm:.2f}€  |   JLCPCB: {self.tj:.2f}€   |   HYBRID (BEST): {self.hybrid_total:.2f}€</b></h3>

            {notes_html}
        </body>
        </html>
        """

        try:
            # PDF Generation
            with open(path, "wb") as pdf_file:
                pisa_status = pisa.CreatePDF(html_content, dest=pdf_file)
            
            if pisa_status.err:
                QMessageBox.critical(self, "Error", "Error creating PDF")
            else:
                QMessageBox.information(self, "Export", "PDF created successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    # --- STANDARD CRUD METHODS (Create, Read, Update, Delete) ---
    def load_projects(self):
        """Loads the project list from the DB into the sidebar list."""
        self.p_list.clear()
        for p in self.session.query(Project).all():
            it = QListWidgetItem(p.name); it.setData(Qt.ItemDataRole.UserRole, p.id)
            self.p_list.addItem(it)

    def add_project(self):
        n, ok = QInputDialog.getText(self, "New", "Name:"); 
        if ok and n: self.session.add(Project(name=n)); self.session.commit(); self.load_projects()

    def delete_project(self):
        if not self.p_list.currentItem(): return
        if QMessageBox.question(self,"Delete","Are you sure?", QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)==QMessageBox.StandardButton.Yes:
            pid = self.p_list.currentItem().data(Qt.ItemDataRole.UserRole)
            self.session.delete(self.session.get(Project, pid)); self.session.commit()
            self.load_projects(); self.tab.setRowCount(0); self.lbl_t.setText("Select Project")

    def select_project(self, item):
        pid = item.data(Qt.ItemDataRole.UserRole)
        self.current_project = self.session.get(Project, pid)
        self.lbl_t.setText(self.current_project.name)
        self.load_bom()

    def load_bom(self):
        """Populates the central table with components of the selected project."""
        if not self.current_project: return
        self.tab.setRowCount(0)
        comps = self.current_project.components
        self.tab.setRowCount(len(comps))
        for r, c in enumerate(comps): self.render_row(r, c)
        self.apply_filter(self.cmb_f.currentText()); self.calc_total()
        self.tab.resizeRowsToContents()

    def render_row(self, r, c):
        self.tab.setItem(r,0, QTableWidgetItem(str(c.id)))
        self.tab.setItem(r,1, QTableWidgetItem(c.mouser_part_number))
        self.tab.setItem(r,2, QTableWidgetItem(c.jlc_part_number))
        self.tab.setItem(r,3, QTableWidgetItem(c.category))
        self.tab.setItem(r,4, QTableWidgetItem(c.description))
        self.tab.setItem(r,5, QTableWidgetItem(str(c.target_qty)))
        
        mtxt = "N/A"; mcol = "#777"
        if c.last_mouser_stock >= c.target_qty: mtxt = f"📦{c.last_mouser_stock}\n{c.last_mouser_price:.2f}€"; mcol = "#aaddaa"
        elif c.last_mouser_stock > -1: mtxt = f"NO STOCK\n({c.last_mouser_stock})"; mcol = "#ff5555"
        mi = QTableWidgetItem(mtxt); mi.setForeground(QColor(mcol)); self.tab.setItem(r, 6, mi)

        jtxt = "N/A"; jcol = "#777"
        if c.last_jlc_stock >= c.target_qty: jtxt = f"📦{c.last_jlc_stock}\n{c.last_jlc_price:.2f}€"; jcol = "#aaddaa"
        elif c.last_jlc_stock > -1: jtxt = f"NO STOCK\n({c.last_jlc_stock})"; jcol = "#ff5555"
        ji = QTableWidgetItem(jtxt); ji.setForeground(QColor(jcol)); self.tab.setItem(r, 7, ji)

    def apply_filter(self, txt):
        for r in range(self.tab.rowCount()): self.tab.setRowHidden(r, txt != "All" and self.tab.item(r,3).text() != txt)

    def add_component(self):
        if not self.current_project: return
        d = ComponentDialog(self)
        if d.exec():
            data = d.get_data()
            self.session.add(Component(project_id=self.current_project.id, mouser_part_number=data['m_pn'], jlc_part_number=data['j_pn'], description=data['desc'], category=data['cat'], target_qty=data['qty'], backup_part=data['backup']))
            self.session.commit(); self.load_bom()

    def edit_component(self):
        r = self.tab.currentRow()
        if r < 0: return
        c = self.session.get(Component, int(self.tab.item(r,0).text()))
        d = ComponentDialog(self, c)
        if d.exec():
            data = d.get_data()
            c.mouser_part_number = data['m_pn']; c.jlc_part_number = data['j_pn']; c.description = data['desc']
            c.category = data['cat']; c.target_qty = data['qty']; c.backup_part = data['backup']
            self.session.commit(); self.load_bom()

    def del_component(self):
        r = self.tab.currentRow()
        if r >= 0: self.session.delete(self.session.get(Component, int(self.tab.item(r,0).text()))); self.session.commit(); self.load_bom()

    def refresh_prices(self):
        try:
            if not self.current_project: return
            for r in range(self.tab.rowCount()): self.tab.item(r, 6).setText("Updating..."); self.tab.item(r, 7).setText("Updating...")
            for c in self.current_project.components:
                worker = DataUpdater(c.id, c.mouser_part_number, c.jlc_part_number, c.target_qty)
                worker.signals.result.connect(self.update_db_and_ui)
                self.threadpool.start(worker)
        except Exception as e:
            QMessageBox.critical(self, "Error!", str(e))

    def update_db_and_ui(self, data):
        c = self.session.get(Component, data['id'])
        if c:
            c.last_mouser_stock = data['mouser_stock']; c.last_mouser_price = data['mouser_price']
            c.last_jlc_stock = data['jlc_stock']; c.last_jlc_price = data['jlc_price']
            c.last_update = data['timestamp'] # Save refresh timestamp
            self.session.commit()
            for r in range(self.tab.rowCount()):
                if self.tab.item(r,0).text() == str(c.id): self.render_row(r, c); break
        self.calc_total()
        self.tab.resizeRowsToContents()

    def calc_total(self):
        tm = 0.0
        tj = 0.0
        hybrid_total = 0.0
        
        if self.current_project:
            for c in self.current_project.components:
                # Validity check
                m_ok = c.last_mouser_stock >= c.target_qty
                j_ok = c.last_jlc_stock >= c.target_qty
                
                # Individual totals
                if m_ok: tm += c.last_mouser_price
                if j_ok: tj += c.last_jlc_price
                
                # Hybrid Total (Best Price Logic)
                price_m = c.last_mouser_price if m_ok else float('inf')
                price_j = c.last_jlc_price if j_ok else float('inf')
                
                best = min(price_m, price_j)
                if best != float('inf'):
                    hybrid_total += best

        self.tm = tm
        self.tj = tj
        self.hybrid_total = hybrid_total
        self.lbl_stat.setText(f"MOUSER: {tm:.2f}€   |   JLCPCB: {tj:.2f}€   |   ⚡ HYBRID (BEST): {hybrid_total:.2f}€")

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    # SETUP PALETTE (Dark Theme)
    p = QPalette()
    p.setColor(QPalette.ColorRole.Window, QColor(30,30,30)); p.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    p.setColor(QPalette.ColorRole.Base, QColor(25,25,25)); p.setColor(QPalette.ColorRole.AlternateBase, QColor(30,30,30))
    p.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white); p.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    p.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white); p.setColor(QPalette.ColorRole.Button, QColor(30,30,30))
    p.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white); p.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    p.setColor(QPalette.ColorRole.Highlight, QColor(0, 122, 204)); p.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
    app.setPalette(p)

    # --- SETUP DATABASE AND FIRST RUN ---
    settings = QSettings("MySoft", "BOMManager")
    db_path = settings.value("db_path", "")
    
    # If no DB path is saved
    if not db_path:
        # Create a custom QMessageBox for initial choice
        msg = QMessageBox()
        msg.setWindowTitle("BOM Manager - Initial Setup")
        msg.setText("Welcome! No database configured.")
        msg.setInformativeText("Do you want to create a new file or open an existing one (e.g., from Cloud)?")
        
        # Add custom buttons
        btn_new = msg.addButton("✨ Create New", QMessageBox.ButtonRole.ActionRole)
        btn_open = msg.addButton("📂 Open Existing", QMessageBox.ButtonRole.ActionRole)
        btn_cancel = msg.addButton("Exit", QMessageBox.ButtonRole.RejectRole)
        
        msg.exec()

        if msg.clickedButton() == btn_cancel:
            sys.exit(0)
        
        elif msg.clickedButton() == btn_new:
            # SAVE FILE NAME (Create New)
            file_path, _ = QFileDialog.getSaveFileName(None, "Create New Database", "my_bom_manager.db", "SQLite Database (*.db)")
            if file_path:
                settings.setValue("db_path", file_path)
                db_path = file_path
                with open(db_path, 'w'): pass # Creates file
            else:
                sys.exit(0)

        elif msg.clickedButton() == btn_open:
            # OPEN FILE NAME (Open Existing - Safe, no overwrite)
            file_path, _ = QFileDialog.getOpenFileName(None, "Select Existing Database", "", "SQLite Database (*.db)")
            if file_path:
                settings.setValue("db_path", file_path)
                db_path = file_path
            else:
                sys.exit(0)

    # Check file existence before loading (if it was deleted manually)
    if not os.path.exists(db_path):
        QMessageBox.critical(None, "Database Error", f"The database file does not exist:\n{db_path}\n\nRestart and select 'Create New' or 'Open Existing'.")
        settings.remove("db_path") # Reset setting
        sys.exit(1)

    # Initialize DB and Session
    session_factory = init_db(db_path)
    w = MainWindow(session_factory)
    w.showMaximized()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
