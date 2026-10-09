#!/usr/bin/env python3
import json
import logging
import os
from pathlib import Path
import shutil
import sys
import traceback

# --- PARCHE CORREGIDO PARA PLATAFORMAS DE PYQT6 ---
if getattr(sys, "frozen", False):
  base_path = sys._MEIPASS
  # Buscamos la carpeta platforms exacta donde vive qwindows.dll
  possible_paths = [
      os.path.join(base_path, "PyQt6", "Qt6", "plugins", "platforms"),
      os.path.join(base_path, "PyQt6", "plugins", "platforms"),
      os.path.join(base_path, " PyQt6", "Qt", "plugins", "platforms"),
  ]
  for p in possible_paths:
    if os.path.exists(p):
      os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = p
      break

print("1. Iniciando script...")

log_dir = Path.home() / "AppData" / "Roaming" / "SinergiaPortableManager"
log_dir.mkdir(parents=True, exist_ok=True)
log_file = log_dir / "sinergia.log"

logging.basicConfig(
    filename=str(log_file),
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)

print("3. Importando PyQt6...")
from PyQt6.QtCore import QSize, Qt
from PyQt6.QtGui import QColor, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)
print("4. PyQt6 importado correctamente.")

print("5. Importando librerías de Windows...")
try:
  import winshell
  from win32com.client import Dispatch
  import win32ui
  import win32gui
  import win32con
  import win32api
  HAS_WIN32 = True
  print("6. Librerías win32/winshell cargadas con éxito.")
except Exception as e:
  HAS_WIN32 = False
  print(f"6. Advertencia: No se pudieron cargar librerías win32: {e}")


class DropZoneWidget(QFrame):
  def __init__(self, main_window):
    super().__init__()
    self.main_window = main_window
    self.setAcceptDrops(True)
    self.setObjectName("dropZone")
    self.setMinimumHeight(75)

    layout = QVBoxLayout()
    self.label = QLabel("📥  Arrastra nuevos .exe portables aquí\n(o haz clic para seleccionarlos)")
    self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    layout.addWidget(self.label)
    self.setLayout(layout)

    self.setStyleSheet("""
        QFrame#dropZone {
            background-color: #1f2335;
            border: 2px dashed #3b4261;
            border-radius: 12px;
        }
        QFrame#dropZone:hover {
            border-color: #7aa2f7;
            background-color: #24283b;
        }
        QLabel {
            color: #c0caf5;
            font-size: 13px;
            font-weight: bold;
            background: transparent;
        }
    """)

  def mousePressEvent(self, event):
    if event.button() == Qt.MouseButton.LeftButton:
      self.main_window.add_portable_dialog()

  def dragEnterEvent(self, event):
    if event.mimeData().hasUrls():
      event.acceptProposedAction()
    else:
      event.ignore()

  def dropEvent(self, event):
    for url in event.mimeData().urls():
      file_path = url.toLocalFile()
      if file_path.lower().endswith(".exe"):
        self.main_window.process_and_add_portable(file_path)
    event.acceptProposedAction()


class SinergiaPortableManager(QMainWindow):
  def __init__(self):
    print("7. Inicializando la ventana principal (QMainWindow)...")
    super().__init__()
    self.setWindowTitle("Sinergia Portable Apps Manager (Windows)")
    self.resize(650, 520)
    self.portables = []

    self.config_dir = log_dir
    self.config_file = self.config_dir / "apps.json"

    self.icon_cache_dir = Path.home() / "AppData" / "Local" / "SinergiaPortableManager" / "icons"
    self.icon_cache_dir.mkdir(parents=True, exist_ok=True)

    self.managed_dir = Path.home() / "PortableApps"
    self.managed_dir.mkdir(parents=True, exist_ok=True)

    self.setStyleSheet("""
        QMainWindow { background-color: #16161e; }
        QPushButton {
            background-color: #24283b; color: #c0caf5;
            border: 1px solid #292e42; border-radius: 8px;
            padding: 10px 16px; font-weight: bold; font-size: 13px;
        }
        QPushButton:hover { background-color: #292e42; border-color: #3b4261; }
        QPushButton#btn_run { background-color: #3d59a1; color: white; border: none; }
        QPushButton#btn_run:hover { background-color: #4c6fc1; }
        QPushButton#btn_remove { background-color: #f7768e; color: white; border: none; }
        QPushButton#btn_remove:hover { background-color: #db4b4b; }
        QListWidget {
            background-color: #1f2335; border: 1px solid #292e42;
            border-radius: 12px; font-size: 13px; color: #c0caf5; padding: 5px;
        }
        QListWidget::item { background-color: #24283b; margin: 4px 6px; padding: 10px; border-radius: 8px; }
        QListWidget::item:selected { background-color: #3d59a1; color: #ffffff; }
    """)

    layout = QVBoxLayout()
    layout.setContentsMargins(20, 20, 20, 20)
    layout.setSpacing(15)

    self.drop_zone = DropZoneWidget(self)
    layout.addWidget(self.drop_zone)

    self.list_widget = QListWidget()
    self.list_widget.setIconSize(QSize(36, 36))
    layout.addWidget(self.list_widget)

    btn_layout = QHBoxLayout()
    self.btn_add = QPushButton("Añadir Ejecutable")
    self.btn_remove = QPushButton("Quitar de la lista")
    self.btn_remove.setObjectName("btn_remove")
    self.btn_run = QPushButton("Ejecutar Seleccionada")
    self.btn_run.setObjectName("btn_run")

    self.btn_add.clicked.connect(self.add_portable_dialog)
    self.btn_run.clicked.connect(self.run_portable)
    self.btn_remove.clicked.connect(self.remove_portable)

    btn_layout.addWidget(self.btn_add)
    btn_layout.addWidget(self.btn_remove)
    btn_layout.addWidget(self.btn_run)
    layout.addLayout(btn_layout)

    container = QWidget()
    container.setLayout(layout)
    self.setCentralWidget(container)

    print("8. Cargando aplicaciones guardadas...")
    self.load_saved_apps()
    print("9. Ventana inicializada con éxito.")

  def load_saved_apps(self):
    if self.config_file.exists():
      try:
        with open(self.config_file, "r", encoding="utf-8") as f:
          saved_paths = json.load(f)
          for path_str in saved_paths:
            if os.path.exists(path_str):
              self.add_app_to_ui(path_str, save=False)
      except Exception as e:
        print(f"Error al cargar apps: {e}")

  def save_apps(self):
    try:
      with open(self.config_file, "w", encoding="utf-8") as f:
        json.dump(self.portables, f, indent=4)
    except Exception as e:
      print(f"Error al guardar apps: {e}")

  def extract_ico_from_exe(self, exe_path):
    if not HAS_WIN32:
      return None
    try:
      ico_path = self.icon_cache_dir / f"{Path(exe_path).stem}.ico"
      ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
      ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)

      large, small = win32gui.ExtractIconEx(exe_path, 0)
      if large:
        win32gui.DestroyIcon(small[0])
        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
        hdc = hdc.CreateCompatibleDC()
        hdc.SelectObject(hbmp)
        hdc.DrawIcon((0, 0), large[0])
        hbmp.SaveBitmapFile(hdc, str(ico_path))
        win32gui.DestroyIcon(large[0])
        return str(ico_path)
    except Exception as e:
      pass
    return None

  def create_windows_shortcut(self, exe_path):
    if not HAS_WIN32:
      return
    try:
      start_menu = winshell.programs()
      sinergia_menu = os.path.join(start_menu, "Sinergia Apps")
      if not os.path.exists(sinergia_menu):
        os.makedirs(sinergia_menu)
      shortcut_path = os.path.join(sinergia_menu, f"{Path(exe_path).stem}.lnk")
      shell = Dispatch('WScript.Shell')
      shortcut = shell.CreateShortcut(shortcut_path)
      shortcut.TargetPath = exe_path
      shortcut.WorkingDirectory = str(Path(exe_path).parent)
      shortcut.save()
    except Exception as e:
      print(f"Error al crear acceso directo: {e}")

  def add_app_to_ui(self, file_path, save=True):
    path = Path(file_path)
    if str(path) in self.portables:
      return
    try:
      icon_path = self.extract_ico_from_exe(str(path))
      item = QListWidgetItem(path.stem)
      if icon_path and os.path.exists(icon_path):
        item.setIcon(QIcon(icon_path))
      self.list_widget.addItem(item)
      self.portables.append(str(path))
      if save:
        self.save_apps()
    except Exception as e:
      print(f"Error al añadir app a UI: {e}")

  def process_and_add_portable(self, file_path):
    source_path = Path(file_path)
    if not source_path.is_relative_to(self.managed_dir):
      target_path = self.managed_dir / source_path.name
      try:
        shutil.copy2(str(source_path), str(target_path))
        file_path = str(target_path)
      except Exception as e:
        print(f"Error al copiar app: {e}")
    self.add_app_to_ui(file_path, save=True)
    self.create_windows_shortcut(file_path)

  def add_portable_dialog(self):
    file_path, _ = QFileDialog.getOpenFileName(
        self, "Seleccionar Ejecutable Portable", "", "Ejecutables (*.exe);;Todos (*)"
    )
    if file_path:
      self.process_and_add_portable(file_path)

  def remove_portable(self):
    selected_items = self.list_widget.selectedItems()
    if not selected_items:
      QMessageBox.warning(self, "Atención", "Selecciona una aplicación de la lista.")
      return
    index = self.list_widget.row(selected_items[0])
    self.list_widget.takeItem(index)
    self.portables.pop(index)
    self.save_apps()

  def run_portable(self):
    selected_items = self.list_widget.selectedItems()
    if not selected_items:
      QMessageBox.warning(self, "Atención", "Selecciona una aplicación para ejecutar.")
      return
    index = self.list_widget.row(selected_items[0])
    app_path = self.portables[index]
    try:
      os.startfile(app_path)
    except Exception as e:
      QMessageBox.critical(self, "Error", f"No se pudo ejecutar: {e}")


if __name__ == "__main__":
  try:
    print("10. Creando instancia de QApplication...")
    app = QApplication(sys.argv)
    print("11. QApplication creada. Creando ventana SinergiaPortableManager...")
    window = SinergiaPortableManager()
    print("12. Mostrando ventana...")
    window.show()
    print("13. Entrando en el bucle principal de la aplicación (exec)...")
    sys.exit(app.exec())
  except Exception as e:
    print(f"¡EXCEPCIÓN CAPTURADA!: {e}")
    traceback.print_exc()
    input("Presiona Enter para salir...")
    sys.exit(1)