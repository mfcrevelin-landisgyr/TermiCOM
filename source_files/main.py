# -*- coding: utf-8 -*-
import serial, serial.tools.list_ports
import pyperclip
import time
import json
import sys
import re
import os

from platformdirs import user_data_dir
from datetime import datetime

from PyQt6.QtGui import QTextCharFormat, QColor, QFont, QTextCursor, QTextDocument
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QMutex
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QMenu,
    QPushButton, QTextEdit, QComboBox, QLabel, QMessageBox
)

# -------------------------------------------------
# Serial forwarding worker with time/baud-aware buffering
# -------------------------------------------------
class SerialForwarderThread(QThread):
    log_signal = pyqtSignal(str, str)  # (text, color)

    def __init__(self, source_conn, dest_conn, direction, baud, ui_log_mutex: QMutex,
                 gap_chars: int = 3, bits_per_char: int = 10):
        """
        direction: 'AtoB' or 'BtoA'
        baud: integer bps
        gap_chars: how many char-times of silence to consider a frame finished
        bits_per_char: 8N1 => 10 bits typical; adjust if your line config differs
        """
        super().__init__()
        self.source_conn = source_conn
        self.dest_conn = dest_conn
        self.direction = direction
        self.running = False

        # Timing / buffering
        self.baud = max(int(baud), 300)
        self.bits_per_char = max(int(bits_per_char), 9)  # conservative minimum
        self.char_time_ns = int(1e9 * self.bits_per_char / self.baud)
        self.flush_gap_ns = max(self.char_time_ns * max(gap_chars, 1), 100_000)  # ≥ 0.1 ms
        self.poll_interval_s = max(self.char_time_ns / 2 / 1e9, 0.0005)

        self.msg_buffer = bytearray()
        self.msg_start_ns = 0
        self.last_rx_ns = 0

        # Optional shared mutex to avoid simultaneous emits from both threads
        self.ui_log_mutex = ui_log_mutex

    def _flush_frame_if_due(self, now_ns: int):
        """Flush buffered frame if idle gap elapsed."""
        if not self.msg_buffer:
            return
        if (now_ns - self.last_rx_ns) >= self.flush_gap_ns*1.5:
            # Build one consolidated log entry
            hex_data = ' '.join(f'{b:02X}' for b in self.msg_buffer)
            atob = self.direction == 'AtoB'
            tag = '[A->B]' if atob else '[A<-B]'
            color = "#f44250" if atob else "#4285f4"
            timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S.%f")
            prefix = f"{tag} {timestamp} : "

            # Ensure atomic emit wrt the other thread
            if self.ui_log_mutex is not None:
                self.ui_log_mutex.lock()
            try:
                self.log_signal.emit(f"{prefix}{hex_data}\n", color)
            finally:
                if self.ui_log_mutex is not None:
                    self.ui_log_mutex.unlock()

            self.msg_buffer.clear()
            self.msg_start_ns = 0
            self.last_rx_ns = 0

    def _append_rx(self, data: bytes, now_ns: int):
        """Append received bytes into buffer and forward to dest."""
        if not data:
            return
        # Start-of-frame timestamp
        if not self.msg_buffer:
            self.msg_start_ns = now_ns
        self.msg_buffer.extend(data)
        self.last_rx_ns = now_ns

        # Forward to the destination port
        try:
            self.dest_conn.write(data)
        except Exception as e:
            # Log once; keep going
            if self.ui_log_mutex is not None:
                self.ui_log_mutex.lock()
            try:
                self.log_signal.emit(f"[ERROR] write({self.direction}) -> {e}\n", "red")
            finally:
                if self.ui_log_mutex is not None:
                    self.ui_log_mutex.unlock()

    def run(self):
        self.running = True
        while self.running:
            try:
                # Pull whatever is currently waiting
                in_wait = self.source_conn.in_waiting
                if in_wait:
                    data = self.source_conn.read(in_wait)
                    now_ns = time.time_ns()
                    if data:
                        self._append_rx(data, now_ns)
                else:
                    # No new data -> check if a frame is ready to flush
                    self._flush_frame_if_due(time.time_ns())

                # Small poll sleep tuned by baud rate
                time.sleep(self.poll_interval_s)

            except Exception as e:
                # Emit error (rare) and keep trying
                if self.ui_log_mutex is not None:
                    self.ui_log_mutex.lock()
                try:
                    self.log_signal.emit(f"[ERROR] Forwarding error ({self.direction}): {e}\n", "red")
                finally:
                    if self.ui_log_mutex is not None:
                        self.ui_log_mutex.unlock()
                time.sleep(0.2)

        # Thread is stopping: flush any remaining buffered frame
        try:
            self._flush_frame_if_due(time.time_ns() + self.flush_gap_ns + 1)
        except Exception:
            pass

    def stop(self):
        self.running = False


# -------------------------------------------------
# Read-only text box with copy/clear context menu
# -------------------------------------------------
class CTextEdit(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        # Read-only, but allow selection by mouse/keyboard
        self.setReadOnly(True)
        self.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
            | Qt.TextInteractionFlag.TextSelectableByKeyboard
        )

        # --- highlight state ---
        self._active_selection = ""
        self._hl_format = QTextCharFormat()
        # Soft yellow highlight; tweak to taste
        self._hl_format.setBackground(QColor('#fff59d'))

        # Re-run highlights when the user changes selection or text changes
        self.selectionChanged.connect(self._on_selection_changed)
        self.document().contentsChange.connect(self._on_contents_change)

    # ----- context menu (unchanged items + improved copy) -----
    def contextMenuEvent(self, event):
        menu = QMenu(self)

        clear_action = menu.addAction("Clear")
        flag_action  = menu.addAction("Flag")
        copy_action  = menu.addAction("Copy")

        action = menu.exec(event.globalPos())

        if action == copy_action:
            # Copy selection if present; otherwise copy entire text
            txt = self.textCursor().selectedText() or self.toPlainText()
            # Normalize paragraph separator to newline for clipboard
            pyperclip.copy(txt.replace('\u2029', '\n'))

        if action == flag_action:
            cursor = self.textCursor()
            fmt = QTextCharFormat()
            fmt.setForeground(QColor('#fcfc4b'))
            # fmt.setFontWeight(QFont.Weight.Bold)
            cursor.setCharFormat(fmt)
            cursor.insertText('◉\n')
            self.setTextCursor(cursor)

        elif action == clear_action:
            self.clear()

    # ----- keyboard shortcuts -----
    def keyPressEvent(self, event):
        # Ctrl+W to clear
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_W:
            self.clear()
        else:
            super().keyPressEvent(event)

    # ----- selection/highlight logic -----
    def _on_selection_changed(self):
        sel = self.textCursor().selectedText()
        # Normalize newlines from selection (QTextEdit uses U+2029 in selections)
        if sel:
            sel = sel.replace('\u2029', '\n')
        # Ignore empty/whitespace-only selections
        self._active_selection = sel if (sel and sel.strip()) else ""
        self._apply_highlights()

    def _on_contents_change(self, *_):
        # If a pattern is active, re-apply highlights when text changes
        if self._active_selection:
            self._apply_highlights()

    def _apply_highlights(self):
        if not self._active_selection:
            self.setExtraSelections([])
            return

        doc = self.document()
        cur = QTextCursor(doc)
        extras = []

        # Case-sensitive search (matches exactly what the user selected)
        find_flags = QTextDocument.FindFlag.FindCaseSensitively

        while True:
            cur = doc.find(self._active_selection, cur, find_flags)
            if cur.isNull():
                break
            sel = QTextEdit.ExtraSelection()
            sel.cursor = cur
            sel.format = self._hl_format
            extras.append(sel)

        self.setExtraSelections(extras)

    def reapply_highlight(self):
        """Helper to re-apply current highlight after programmatic edits."""
        if self._active_selection:
            self._apply_highlights()
        else:
            self.setExtraSelections([])




# -------------------------------------------------
# Main Application
# -------------------------------------------------
class SerialMiddlemanApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WolfWire - COM Port Middleman")
        self.CACHE_FILE = os.path.join(user_data_dir("WolfWire", "WolfWire"), "com_cache.json")

        self.serial_A = None
        self.serial_B = None
        self.thread_AtoB = None
        self.thread_BtoA = None

        # guard to avoid recursive updates when filtering combos
        self._updating_combos = False

        # Shared UI log mutex (prevents simultaneous emit from both threads)
        self.log_mutex = QMutex()

        self.init_ui()
        self.refresh_com_ports()

        self._restore_cached_selection()

    # ---------- Cache ----------
    def _restore_cached_selection(self):
        try:
            with open(self.CACHE_FILE, "r", encoding="utf-8") as f:
                self.cached_ports = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.cached_ports = {}

        if not self.cached_ports:
            return

        ports = self._list_active_ports()
        rates = [self.baud_rate_combo.itemText(i) for i in range(self.baud_rate_combo.count())]

        portA = self.cached_ports.get("portA")
        portB = self.cached_ports.get("portB")
        baud  = self.cached_ports.get("baud")

        if portA in ports:
            self.com_port_combo_A.setCurrentText(portA)
        if portB in ports:
            self.com_port_combo_B.setCurrentText(portB)
        if baud in rates:
            self.baud_rate_combo.setCurrentText(baud)

    def _save_cached_ports(self):
        os.makedirs(os.path.dirname(self.CACHE_FILE), exist_ok=True)
        data = {
            "portA": self.com_port_combo_A.currentText(),
            "portB": self.com_port_combo_B.currentText(),
            "baud": self.baud_rate_combo.currentText()
        }
        with open(self.CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)

    # ---------- UI ----------
    def init_ui(self):
        layout = QVBoxLayout()
        config_layout = QHBoxLayout()

        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.clicked.connect(self.refresh_com_ports)

        self.com_port_combo_A = QComboBox()
        self.com_port_combo_B = QComboBox()
        self.com_port_combo_A.currentTextChanged.connect(self._on_portA_selected)
        self.com_port_combo_B.currentTextChanged.connect(self._on_portB_selected)

        self.baud_rate_combo = QComboBox()
        self.baud_rate_combo.addItems([
                "300", "600", "1200", "2400", "4800", "9600", "14400", 
                "19200", "28800", "38400", "57600", "115200", "128000", "256000"
            ])
        self.baud_rate_combo.setCurrentText("115200")
        self.baud_rate_combo.currentTextChanged.connect(self._on_baud_selected)

        config_layout.addWidget(self.refresh_button)
        config_layout.addWidget(QLabel("COM Port A:"))
        config_layout.addWidget(self.com_port_combo_A)
        config_layout.addWidget(QLabel("COM Port B:"))
        config_layout.addWidget(self.com_port_combo_B)
        config_layout.addWidget(QLabel("Baud Rate:"))
        config_layout.addWidget(self.baud_rate_combo)
        config_layout.addStretch()

        self.connect_button = QPushButton("Connect")
        self.connect_button.clicked.connect(self.toggle_connection)

        self.status_label = QLabel("Disconnected")
        self.status_label.setStyleSheet("background-color: red;")
        config_layout.addWidget(self.connect_button)
        config_layout.addWidget(self.status_label)

        layout.addLayout(config_layout)

        self.textbox = CTextEdit()
        layout.addWidget(self.textbox)

        self.setLayout(layout)

    # ---------- Port listing & filtering ----------
    def _list_active_ports(self):
        """Return list of strings like ['COM3', 'COM4', ...] on Windows or '/dev/tty...' on *nix."""
        return [port.device for port in serial.tools.list_ports.comports()]

    def _fill_combo(self, combo: QComboBox, items, keep_selection=None):
        """Fill combo with items, optionally keeping previous selection if still present."""
        combo.blockSignals(True)
        current = keep_selection if keep_selection is not None else combo.currentText()
        combo.clear()
        combo.addItems(items)
        if current and current in items:
            combo.setCurrentText(current)
        combo.blockSignals(False)

    def refresh_com_ports(self):
        """
        Refresh both combos with active ports and keep them mutually exclusive.
        """
        if self._updating_combos:
            return
        self._updating_combos = True

        ports = self._list_active_ports()

        # Remember current selections to try to preserve them
        selA = self.com_port_combo_A.currentText()
        selB = self.com_port_combo_B.currentText()

        # First, fill both with the full list
        self._fill_combo(self.com_port_combo_A, ports, keep_selection=selA)
        self._fill_combo(self.com_port_combo_B, ports, keep_selection=selB)

        # Disable Connect if less than 2 ports are available
        self.connect_button.setEnabled(len(ports) >= 2)

        self._updating_combos = False

    def _on_portA_selected(self, text):
        self._save_cached_ports()

    def _on_portB_selected(self, text):
        self._save_cached_ports()


    def _on_baud_selected(self, text):
        self._save_cached_ports()

    # ---------- Connect / Disconnect ----------
    def toggle_connection(self):
        if self.serial_A and self.serial_A.is_open:
            self.disconnect_ports()
        else:
            self.connect_ports()

    def connect_ports(self):
        ports = self._list_active_ports()
        if len(ports) < 2:
            self.show_error("Not enough active COM ports detected (need at least 2).")
            return

        portA = self.com_port_combo_A.currentText()
        portB = self.com_port_combo_B.currentText()
        if not portA or not portB:
            self.show_error("Please select both COM Port A and COM Port B.")
            return
        if portA == portB:
            self.show_error("COM Port A and COM Port B must be different.")
            return

        baud_rate = int(self.baud_rate_combo.currentText())

        # Open ports
        try:
            self.serial_A = serial.Serial(portA, baud_rate, timeout=1)
            self.serial_B = serial.Serial(portB, baud_rate, timeout=1)
        except serial.SerialException as e:
            self.show_error(f"Error opening ports:\n{e}")
            # Cleanup if one opened successfully
            try:
                if self.serial_A and self.serial_A.is_open:
                    self.serial_A.close()
            except Exception:
                pass
            try:
                if self.serial_B and self.serial_B.is_open:
                    self.serial_B.close()
            except Exception:
                pass
            self.serial_A = None
            self.serial_B = None
            return

        # Start forwarders (with buffering)
        # bits_per_char: 10 for 8N1; adjust if you later expose parity/stop settings
        self.thread_AtoB = SerialForwarderThread(self.serial_A, self.serial_B, "AtoB",
                                                 baud_rate, self.log_mutex,
                                                 gap_chars=3, bits_per_char=10)
        self.thread_BtoA = SerialForwarderThread(self.serial_B, self.serial_A, "BtoA",
                                                 baud_rate, self.log_mutex,
                                                 gap_chars=3, bits_per_char=10)

        self.thread_AtoB.log_signal.connect(self.log_text)
        self.thread_BtoA.log_signal.connect(self.log_text)
        self.thread_AtoB.start()
        self.thread_BtoA.start()

        # UI state
        self.status_label.setText("Connected")
        self.status_label.setStyleSheet("background-color: green;")
        self.connect_button.setText("Disconnect")
        self.com_port_combo_A.setEnabled(False)
        self.com_port_combo_B.setEnabled(False)
        self.baud_rate_combo.setEnabled(False)
        self.refresh_button.setEnabled(False)
        self.log_text(f"[INFO] Bridging {portA} <=> {portB} @ {baud_rate} bps\n", "green")

    def disconnect_ports(self):
        # Stop threads
        if self.thread_AtoB:
            self.thread_AtoB.stop()
            self.thread_AtoB.wait()
            self.thread_AtoB = None
        if self.thread_BtoA:
            self.thread_BtoA.stop()
            self.thread_BtoA.wait()
            self.thread_BtoA = None

        # Close serials
        if self.serial_A and self.serial_A.is_open:
            self.serial_A.close()
        if self.serial_B and self.serial_B.is_open:
            self.serial_B.close()
        self.serial_A = None
        self.serial_B = None

        # UI back to normal
        self.status_label.setText("Disconnected")
        self.status_label.setStyleSheet("background-color: red;")
        self.connect_button.setText("Connect")
        self.com_port_combo_A.setEnabled(True)
        self.com_port_combo_B.setEnabled(True)
        self.baud_rate_combo.setEnabled(True)
        self.refresh_button.setEnabled(True)

        # Refresh available ports (in case a device was unplugged/plugged)
        self.refresh_com_ports()

    # ---------- Logging & helpers ----------
    def _log_text_internal(self, text, color="gray"):
        cursor = self.textbox.textCursor()
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(color))
        # fmt.setFontWeight(QFont.Weight.Bold)
        cursor.setCharFormat(fmt)
        cursor.insertText(text)
        self.textbox.setTextCursor(cursor)

    def log_text(self, text, color):
        self._log_text_internal(text, color)

    def show_error(self, message):
        QMessageBox.warning(self, "Error", message)

    def closeEvent(self, event):
        self.disconnect_ports()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SerialMiddlemanApp()
    window.show()
    sys.exit(app.exec())