#!/usr/bin/env python3
"""
Lightweight PDF Reader & Converter

Features:
- Open PDFs via menu or drag-and-drop
- Smooth scrolling, Next/Previous page, Zoom In/Out
- Highlight and Underline annotations (saved back to PDF)
- Convert PDF -> DOCX (using pdf2docx)
- Convert DOCX -> PDF (using LibreOffice headless, or docx2pdf fallback)
- 100% local, no background services or internet

Setup:
1. Install Python 3.9+
2. Install dependencies:
   pip install PyQt6 PyMuPDF pdf2docx
   Optional for DOCX->PDF:
   - Install LibreOffice (https://www.libreoffice.org/) and ensure 'libreoffice' or 'soffice' is in PATH
   - OR install docx2pdf (Windows only, requires MS Word): pip install docx2pdf
3. Run:
   python pdf_reader.py
"""

import sys
import os
import shutil
import tempfile
import subprocess
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QScrollArea,
    QToolBar, QFileDialog, QMessageBox, QProgressDialog
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QRect
from PyQt6.QtGui import QImage, QPixmap, QPainter, QColor, QPen, QBrush, QAction, QKeySequence

import fitz  # PyMuPDF
from pdf2docx import Converter


class PageWidget(QWidget):
    """Widget that displays a single PDF page and handles annotation selection."""

    def __init__(self, page_index, selection_callback, parent=None):
        super().__init__(parent)
        self.page_index = page_index
        self.selection_callback = selection_callback
        self.pixmap = QPixmap()
        self.annotation_mode = None
        self.selecting = False
        self.selection_start = None
        self.selection_rect = None
        self.setMouseTracking(False)

    def set_pixmap(self, pixmap: QPixmap):
        """Update the page image and resize widget accordingly."""
        self.pixmap = pixmap
        self.setFixedSize(pixmap.size())
        self.update()

    def set_annotation_mode(self, mode: str | None):
        """Set annotation mode ('highlight', 'underline', or None)."""
        self.annotation_mode = mode
        if mode:
            self.setCursor(Qt.CursorShape.CrossCursor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        if not self.pixmap.isNull():
            painter.drawPixmap(0, 0, self.pixmap)

        # Draw selection rectangle overlay
        if self.selecting and self.selection_rect:
            painter.setPen(QPen(QColor(0, 120, 215), 1, Qt.PenStyle.DashLine))
            painter.setBrush(QBrush(QColor(0, 120, 215, 60)))
            painter.drawRect(self.selection_rect)
        painter.end()

    def mousePressEvent(self, event):
        if self.annotation_mode and event.button() == Qt.MouseButton.LeftButton:
            self.selecting = True
            self.selection_start = event.position().toPoint()
            self.selection_rect = QRect(self.selection_start, self.selection_start)
            self.update()

    def mouseMoveEvent(self, event):
        if self.selecting:
            current = event.position().toPoint()
            self.selection_rect = QRect(self.selection_start, current).normalized()
            self.update()

    def mouseReleaseEvent(self, event):
        if self.selecting and event.button() == Qt.MouseButton.LeftButton:
            self.selecting = False
            rect = self.selection_rect
            self.selection_rect = None
            self.update()
            if rect and not rect.isEmpty():
                # Notify MainWindow about the selected rectangle
                self.selection_callback(self.page_index, rect)


class ConversionThread(QThread):
    """Worker thread for long-running conversions."""
    finished_ok = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, func, *args):
        super().__init__()
        self.func = func
        self.args = args

    def run(self):
        try:
            result = self.func(*self.args)
            self.finished_ok.emit(result or "")
        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Lightweight PDF Reader")
        self.setGeometry(100, 100, 900, 1000)

        self.doc = None
        self.file_path = None
        self.zoom_factor = 1.0
        self.page_widgets = []
        self.current_page = 0
        self.annotation_mode = None

        self.init_ui()
        self.setAcceptDrops(True)

    # ----------------------------------------------------------------------
    # UI Setup
    # ----------------------------------------------------------------------
    def init_ui(self):
        # Central scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.container = QWidget()
        self.page_layout = QVBoxLayout(self.container)
        self.page_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.page_layout.setContentsMargins(0, 0, 0, 0)
        self.page_layout.setSpacing(5)
        self.scroll_area.setWidget(self.container)
        self.setCentralWidget(self.scroll_area)

        # Main toolbar
        toolbar = self.addToolBar("Main")
        open_action = QAction("Open PDF", self)
        open_action.setShortcut(QKeySequence("Ctrl+O"))
        open_action.triggered.connect(self.open_pdf)
        toolbar.addAction(open_action)

        save_action = QAction("Save Annotated PDF", self)
        save_action.setShortcut(QKeySequence("Ctrl+S"))
        save_action.triggered.connect(self.save_annotated)
        toolbar.addAction(save_action)

        toolbar.addSeparator()
        prev_action = QAction("Previous Page", self)
        prev_action.setShortcut(QKeySequence("PgUp"))
        prev_action.triggered.connect(self.prev_page)
        toolbar.addAction(prev_action)

        next_action = QAction("Next Page", self)
        next_action.setShortcut(QKeySequence("PgDown"))
        next_action.triggered.connect(self.next_page)
        toolbar.addAction(next_action)

        toolbar.addSeparator()
        zoom_in_action = QAction("Zoom In", self)
        zoom_in_action.setShortcut(QKeySequence("Ctrl++"))
        zoom_in_action.triggered.connect(self.zoom_in)
        toolbar.addAction(zoom_in_action)

        zoom_out_action = QAction("Zoom Out", self)
        zoom_out_action.setShortcut(QKeySequence("Ctrl+-"))
        zoom_out_action.triggered.connect(self.zoom_out)
        toolbar.addAction(zoom_out_action)

        toolbar.addSeparator()
        pdf_to_docx_action = QAction("PDF to DOCX", self)
        pdf_to_docx_action.triggered.connect(self.convert_pdf_to_docx)
        toolbar.addAction(pdf_to_docx_action)

        docx_to_pdf_action = QAction("DOCX to PDF", self)
        docx_to_pdf_action.triggered.connect(self.convert_docx_to_pdf)
        toolbar.addAction(docx_to_pdf_action)

        # Annotation toolbar
        annotation_toolbar = self.addToolBar("Annotations")
        self.highlight_action = QAction("Highlight", self, checkable=True)
        self.highlight_action.triggered.connect(
            lambda checked: self.set_annotation_mode("highlight" if checked else None)
        )
        annotation_toolbar.addAction(self.highlight_action)

        self.underline_action = QAction("Underline", self, checkable=True)
        self.underline_action.triggered.connect(
            lambda checked: self.set_annotation_mode("underline" if checked else None)
        )
        annotation_toolbar.addAction(self.underline_action)

        self.statusBar().showMessage("Open a PDF to begin")

    # ----------------------------------------------------------------------
    # Annotation Mode
    # ----------------------------------------------------------------------
    def set_annotation_mode(self, mode: str | None):
        self.annotation_mode = mode
        # Ensure only one annotation tool is checked at a time
        self.highlight_action.setChecked(mode == "highlight")
        self.underline_action.setChecked(mode == "underline")
        for pw in self.page_widgets:
            pw.set_annotation_mode(mode)

    # ----------------------------------------------------------------------
    # PDF Loading & Rendering
    # ----------------------------------------------------------------------
    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if path:
            self.load_pdf(path)

    def load_pdf(self, path: str):
        try:
            if self.doc:
                self.doc.close()
            self.doc = fitz.open(path)
            self.file_path = path
            self.setWindowTitle(f"Lightweight PDF Reader - {os.path.basename(path)}")
            self.zoom_factor = 1.0
            self.current_page = 0
            self.render_all_pages()
            self.statusBar().showMessage(f"Loaded {path} ({len(self.doc)} pages)")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open PDF: {e}")

    def render_all_pages(self):
        """Clear and re-render all pages at current zoom."""
        while self.page_layout.count():
            item = self.page_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.page_widgets.clear()

        if not self.doc:
            return

        for i in range(len(self.doc)):
            pw = PageWidget(i, self.apply_annotation_from_selection)
            self.page_layout.addWidget(pw)
            self.page_widgets.append(pw)

        for pw in self.page_widgets:
            self.render_page(pw.page_index)
            pw.set_annotation_mode(self.annotation_mode)

        if self.page_widgets:
            self.scroll_to_page(self.current_page)

    def render_page(self, page_index: int):
        """Render a single page to QPixmap and update its widget."""
        if not self.doc or page_index < 0 or page_index >= len(self.page_widgets):
            return
        page = self.doc[page_index]
        matrix = fitz.Matrix(self.zoom_factor, self.zoom_factor)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
        pixmap = QPixmap.fromImage(img)
        self.page_widgets[page_index].set_pixmap(pixmap)

    def scroll_to_page(self, page_index: int):
        if page_index < 0 or page_index >= len(self.page_widgets):
            return
        widget = self.page_widgets[page_index]
        self.scroll_area.ensureWidgetVisible(widget, 0, 0)
        self.current_page = page_index
        self.statusBar().showMessage(f"Page {page_index + 1} of {len(self.doc)}")

    # ----------------------------------------------------------------------
    # Navigation & Zoom
    # ----------------------------------------------------------------------
    def next_page(self):
        if self.doc and self.current_page < len(self.doc) - 1:
            self.scroll_to_page(self.current_page + 1)

    def prev_page(self):
        if self.doc and self.current_page > 0:
            self.scroll_to_page(self.current_page - 1)

    def zoom_in(self):
        self.zoom_factor = min(self.zoom_factor * 1.25, 5.0)
        self.render_all_pages()

    def zoom_out(self):
        self.zoom_factor = max(self.zoom_factor * 0.8, 0.2)
        self.render_all_pages()

    # ----------------------------------------------------------------------
    # Annotation Handling
    # ----------------------------------------------------------------------
    def apply_annotation_from_selection(self, page_index: int, widget_rect: QRect):
        """Map selection rectangle to PDF coordinates and apply highlight/underline."""
        if not self.doc or self.annotation_mode not in ("highlight", "underline"):
            return

        page = self.doc[page_index]
        zoom = self.zoom_factor

        # Convert widget pixel rect to PDF points
        sel_rect = fitz.Rect(
            widget_rect.x() / zoom,
            widget_rect.y() / zoom,
            (widget_rect.x() + widget_rect.width()) / zoom,
            (widget_rect.y() + widget_rect.height()) / zoom
        )

        words = page.get_text("words")
        selected_quads = []
        for w in words:
            word_rect = fitz.Rect(w[0], w[1], w[2], w[3])
            if word_rect.intersects(sel_rect):
                selected_quads.append(fitz.Quad(word_rect))

        if not selected_quads:
            QMessageBox.information(self, "No Text", "No text found in the selected area.")
            return

        if self.annotation_mode == "highlight":
            annot = page.add_highlight_annot(selected_quads)
        else:
            annot = page.add_underline_annot(selected_quads)

        annot.set_info(title="Lightweight PDF Reader", content="")
        annot.update()

        # Re-render the page to show the new annotation
        self.render_page(page_index)
        self.statusBar().showMessage(f"Added {self.annotation_mode} to {len(selected_quads)} word(s)")

    def save_annotated(self):
        if not self.doc:
            return
        default_name = os.path.splitext(self.file_path)[0] + "_annotated.pdf" if self.file_path else "annotated.pdf"
        path, _ = QFileDialog.getSaveFileName(self, "Save Annotated PDF", default_name, "PDF Files (*.pdf)")
        if path:
            try:
                self.doc.save(path)
                self.statusBar().showMessage(f"Saved to {path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save: {e}")

    # ----------------------------------------------------------------------
    # Conversion Helpers
    # ----------------------------------------------------------------------
    def run_conversion(self, func, *args, message="Working..."):
        """Run a conversion function in a separate thread with a progress dialog."""
        self.progress = QProgressDialog(message, None, 0, 0, self)
        self.progress.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress.setCancelButton(None)
        self.progress.show()

        self.thread = ConversionThread(func, *args)
        self.thread.finished_ok.connect(self.on_conversion_finished)
        self.thread.error.connect(self.on_conversion_error)
        self.thread.start()

    def on_conversion_finished(self, result: str):
        self.progress.close()
        QMessageBox.information(self, "Success", result or "Conversion completed.")

    def on_conversion_error(self, error: str):
        self.progress.close()
        QMessageBox.critical(self, "Error", f"Conversion failed: {error}")

    # ----------------------------------------------------------------------
    # PDF -> DOCX
    # ----------------------------------------------------------------------
    def convert_pdf_to_docx(self):
        if not self.doc or not self.file_path:
            QMessageBox.warning(self, "No PDF", "Open a PDF first.")
            return
        default_name = os.path.splitext(self.file_path)[0] + ".docx"
        path, _ = QFileDialog.getSaveFileName(self, "Convert PDF to DOCX", default_name, "DOCX Files (*.docx)")
        if path:
            self.run_conversion(self._pdf_to_docx, self.file_path, path, "Converting PDF to DOCX...")

    def _pdf_to_docx(self, pdf_path: str, docx_path: str):
        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()
        return f"PDF converted to DOCX:\n{docx_path}"

    # ----------------------------------------------------------------------
    # DOCX -> PDF
    # ----------------------------------------------------------------------
    def convert_docx_to_pdf(self):
        input_path, _ = QFileDialog.getOpenFileName(self, "Select DOCX", "", "DOCX Files (*.docx)")
        if not input_path:
            return
        output_path, _ = QFileDialog.getSaveFileName(
            self, "Save PDF", os.path.splitext(input_path)[0] + ".pdf", "PDF Files (*.pdf)"
        )
        if output_path:
            self.run_conversion(self._docx_to_pdf, input_path, output_path, "Converting DOCX to PDF...")

    def _docx_to_pdf(self, input_path: str, output_path: str):
        """Convert DOCX to PDF using LibreOffice headless, falling back to docx2pdf."""
        libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
        if libreoffice:
            out_dir = tempfile.mkdtemp()
            subprocess.run(
                [libreoffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, input_path],
                check=True, capture_output=True, timeout=300
            )
            generated = os.path.join(out_dir, os.path.splitext(os.path.basename(input_path))[0] + ".pdf")
            if os.path.exists(generated):
                shutil.move(generated, output_path)
                shutil.rmtree(out_dir, ignore_errors=True)
                return f"DOCX converted to PDF:\n{output_path}"
            else:
                raise RuntimeError("LibreOffice did not produce a PDF file.")
        else:
            try:
                from docx2pdf import convert
                convert(input_path, output_path)
                return f"DOCX converted to PDF:\n{output_path}"
            except ImportError:
                raise RuntimeError(
                    "LibreOffice not found. Install LibreOffice or docx2pdf for DOCX->PDF conversion."
                )

    # ----------------------------------------------------------------------
    # Drag & Drop
    # ----------------------------------------------------------------------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(".pdf"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".pdf"):
                self.load_pdf(path)
                break

    # ----------------------------------------------------------------------
    # Cleanup
    # ----------------------------------------------------------------------
    def closeEvent(self, event):
        if self.doc:
            self.doc.close()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("Lightweight PDF Reader")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
