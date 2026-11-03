from __future__ import annotations

# ============================================================
# DANIEL PDF STUDIO
# Versão revisada — aplicação desktop local
#
# Execução:
#     python daniel_pdf_studio.py
#
# Dependências:
#     python -m pip install --upgrade \
#         PyQt6 PyMuPDF Pillow openpyxl python-pptx ezdxf pdf2docx
#
# Observações:
# - Exit code 0 significa encerramento normal.
# - OCR nativo: somente Windows 10/11.
# - Assinatura inserida: visual, não assinatura digital criptográfica.
# - DOCX/XLSX/PPTX/DXF: conversões aproximadas; revise os resultados.
# ============================================================

import sys
import os
import io
import json
import math
import uuid
import logging
import datetime
import tempfile
import subprocess
import traceback
import functools
from pathlib import Path
from contextlib import contextmanager

APP_NAME = "Daniel PDF Studio"
APP_VERSION = "5.2"

BG = "#080D18"
PANEL = "#101A2B"
CYAN = "#45D9F5"
GREEN = "#38D9A9"
TEXT = "#E6EDF7"
MUTED = "#95A7BE"

COMMENT_TAG = "DPS_COMMENT"
STAMP_TAG = "DPS_STAMP"
SIGNATURE_TAG = "DPS_SIGNATURE"
IMAGE_TAG = "DPS_IMAGE"
RULER_TAG = "DPS_RULER"

ALLOWED_INFO_KEYS = {
    "content", "title", "subject",
    "creationDate", "modDate", "creator"
}

# ============================================================
# LOG
# ============================================================

LOGGER = logging.getLogger("daniel_pdf_studio")
LOGGER.setLevel(logging.INFO)

_console_handler = logging.StreamHandler(sys.stderr)
_console_handler.setFormatter(
    logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
)
LOGGER.addHandler(_console_handler)

LOG_PATH = None

try:
    log_dir = Path.home() / ".daniel_pdf_studio"
    log_dir.mkdir(parents=True, exist_ok=True)
    LOG_PATH = log_dir / "studio.log"

    _file_handler = logging.FileHandler(
        LOG_PATH, encoding="utf-8"
    )
    _file_handler.setFormatter(
        logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    )
    LOGGER.addHandler(_file_handler)
except OSError:
    LOGGER.warning("Não foi possível habilitar o arquivo de log.")

# ============================================================
# DEPENDÊNCIAS OBRIGATÓRIAS
# ============================================================

try:
    from PyQt6.QtCore import (
        Qt, QTimer, QPoint, QPointF, QRect, QRectF,
        QSize, QByteArray, QBuffer, QIODevice,
        QEvent, QMimeData, QThread, pyqtSignal
    )
    from PyQt6.QtGui import (
        QImage, QPixmap, QPainter, QPen, QColor,
        QFont, QIcon, QPainterPath, QDrag,
        QKeySequence, QShortcut, QTextCursor
    )
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QFrame,
        QLabel, QVBoxLayout, QHBoxLayout, QFormLayout,
        QPushButton, QToolButton, QTabWidget, QSplitter,
        QScrollArea, QListWidget, QListWidgetItem,
        QTreeWidget, QTreeWidgetItem, QAbstractItemView,
        QDialog, QFileDialog, QMessageBox, QInputDialog,
        QTextEdit, QLineEdit, QSpinBox, QDoubleSpinBox,
        QComboBox, QCheckBox, QColorDialog,
        QProgressDialog, QDialogButtonBox, QStatusBar
    )
    from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
except ImportError as exc:
    LOGGER.error(
        "PyQt6 indisponível. Instale com: "
        "python -m pip install PyQt6"
    )
    raise SystemExit(1) from exc

try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf

        if not hasattr(pymupdf, "open") or not hasattr(pymupdf, "Rect"):
            raise ImportError("O módulo fitz instalado não é o PyMuPDF.")
    except ImportError as exc:
        LOGGER.error(
            "PyMuPDF indisponível. Instale com: "
            "python -m pip install --upgrade PyMuPDF"
        )
        raise SystemExit(1) from exc

try:
    from PIL import Image
except ImportError as exc:
    LOGGER.error(
        "Pillow indisponível. Instale com: "
        "python -m pip install Pillow"
    )
    raise SystemExit(1) from exc

# openpyxl, python-pptx, ezdxf e pdf2docx são carregados
# somente quando a respectiva conversão for utilizada.

# ============================================================
# UTILITÁRIOS E TRATAMENTO DE ERROS
# ============================================================


def pdf_date():
    now = datetime.datetime.now().astimezone()
    offset = now.strftime("%z")
    return (
        now.strftime("D:%Y%m%d%H%M%S")
        + offset[:3] + "'" + offset[3:] + "'"
    )


def same_path(first, second):
    first = os.path.normcase(os.path.realpath(os.fspath(first)))
    second = os.path.normcase(os.path.realpath(os.fspath(second)))
    return first == second


def safe_filename(name):
    name = str(name or "anexo.bin").replace("\\", "/")
    name = name.rsplit("/", 1)[-1]
    for char in '<>:"|?*\x00':
        name = name.replace(char, "_")
    name = "".join(c for c in name if ord(c) >= 32).rstrip(" .")
    if name in ("", ".", ".."):
        return "anexo.bin"

    stem = name.split(".", 1)[0].upper()
    reserved = {
        "CON", "PRN", "AUX", "NUL",
        *(f"COM{i}" for i in range(1, 10)),
        *(f"LPT{i}" for i in range(1, 10))
    }
    if stem in reserved:
        name = "_" + name
    return name


def safe_set_info(annot, **changes):
    info = {
        key: value
        for key, value in annot.info.items()
        if key in ALLOWED_INFO_KEYS and value is not None
    }
    info.update({
        key: str(value)
        for key, value in changes.items()
        if key in ALLOWED_INFO_KEYS and value is not None
    })
    annot.set_info(info=info)


def annotation_subject(annot):
    return annot.info.get("subject", "") or ""


def is_comment(annot):
    return (
        annot.type[0] == pymupdf.PDF_ANNOT_TEXT
        or annotation_subject(annot).startswith(COMMENT_TAG)
        or annot.info.get("title", "") in {
            "COMMENT_MARKER", "STICKY_NOTE"
        }
    )


def annotation_is_image(annot):
    return (
        annotation_subject(annot).startswith(
            (STAMP_TAG, SIGNATURE_TAG, IMAGE_TAG)
        )
        or annot.info.get("title", "") in {"STAMP", "SIGNATURE"}
    )


def annotation_is_ruler(annot):
    return (
        annotation_subject(annot).startswith(RULER_TAG)
        or annot.info.get("title", "") == "RULER"
    )


def get_annot(page, xref):
    if xref is None:
        return None
    try:
        return page.load_annot(int(xref))
    except (RuntimeError, ValueError):
        return None


def rgb_tuple(color):
    return color.redF(), color.greenF(), color.blueF()


def int_to_rgb(value):
    value = int(value or 0)
    return (
        (value >> 16) & 255,
        (value >> 8) & 255,
        value & 255
    )


def rgb_hex(rgb):
    return "".join(f"{int(value):02X}" for value in rgb)


def escape_html(text):
    return (
        str(text).replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def png_bytes(image):
    if image.isNull():
        raise ValueError("A imagem está vazia ou é inválida.")

    data = QByteArray()
    buffer = QBuffer(data)

    if not buffer.open(QIODevice.OpenModeFlag.WriteOnly):
        raise RuntimeError("Não foi possível criar o buffer PNG.")

    try:
        if not image.save(buffer, "PNG"):
            raise RuntimeError("Não foi possível codificar a imagem.")
    finally:
        buffer.close()

    return bytes(data)


def qimage_from_pixmap(pix):
    if pix.colorspace is None:
        raise ValueError("Pixmap sem espaço de cores.")

    if pix.colorspace.n != 3:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)

    if pix.alpha:
        pix = pymupdf.Pixmap(pix, 0)

    return QImage(
        pix.samples, pix.width, pix.height, pix.stride,
        QImage.Format.Format_RGB888
    ).copy()


def choose_pdf_output(parent, title, suggested):
    path, _ = QFileDialog.getSaveFileName(
        parent, title, str(suggested), "PDF (*.pdf)"
    )
    if path and not path.lower().endswith(".pdf"):
        path += ".pdf"
    return path


def reject_source_output(parent, path, source_path):
    if same_path(path, source_path):
        QMessageBox.warning(
            parent, "Arquivo de saída",
            "Escolha outro nome para não sobrescrever o PDF aberto.\n"
            "Para atualizar o documento atual, utilize Salvar."
        )
        return True
    return False


def native_open(path):
    path = os.path.abspath(path)
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def show_error(parent, title, error):
    LOGGER.error(
        "%s: %s", title, error,
        exc_info=(
            type(error), error, error.__traceback__
        )
    )
    suffix = f"\n\nDiagnóstico registrado em:\n{LOG_PATH}" if LOG_PATH else ""
    QMessageBox.critical(parent, title, str(error) + suffix)


def guarded(function):
    """
    Protege callbacks e slots contra exceções.

    Sinais Qt podem enviar argumentos adicionais, como o booleano
    checked de clicked(bool). Esses argumentos são descartados
    somente quando excedem a capacidade posicional do callback.

    Argumentos necessários, como índices de abas, são preservados.
    Funções que aceitam *args recebem todos os argumentos.
    """
    import inspect

    signature = inspect.signature(function)

    accepts_varargs = any(
        parameter.kind == inspect.Parameter.VAR_POSITIONAL
        for parameter in signature.parameters.values()
    )

    max_positional = sum(
        parameter.kind in (
            inspect.Parameter.POSITIONAL_ONLY,
            inspect.Parameter.POSITIONAL_OR_KEYWORD,
        )
        for parameter in signature.parameters.values()
    )

    @functools.wraps(function)
    def wrapper(*args, **kwargs):
        call_args = (
            args
            if accepts_varargs
            else args[:max_positional]
        )

        try:
            # Valida antes da chamada. A função é executada uma única vez.
            signature.bind(*call_args, **kwargs)
            return function(*call_args, **kwargs)

        except Exception as exc:
            parent = None

            # Métodos decorados diretamente recebem self em args.
            if args and isinstance(args[0], QWidget):
                parent = args[0]

            # Métodos vinculados, como guarded(self.open_pdfs),
            # já possuem self associado à função.
            if parent is None:
                bound_owner = getattr(function, "__self__", None)
                if isinstance(bound_owner, QWidget):
                    parent = bound_owner

            if parent is None:
                parent = QApplication.activeWindow()

            show_error(parent, "Erro de operação", exc)
            return None

    return wrapper


def global_exception_handler(exc_type, exc_value, exc_tb):
    LOGGER.error(
        "Exceção não tratada",
        exc_info=(exc_type, exc_value, exc_tb)
    )
    app = QApplication.instance()
    if app is not None:
        QMessageBox.critical(
            QApplication.activeWindow(),
            "Erro não tratado",
            f"{exc_value}\n\n"
            "A operação não foi concluída. Confira o documento.\n"
            + (f"Log: {LOG_PATH}" if LOG_PATH else "")
        )


sys.excepthook = global_exception_handler
