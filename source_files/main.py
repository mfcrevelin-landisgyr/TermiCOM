import os
import sys
import time
import serial
import serial.tools.list_ports
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QTextEdit, QLineEdit, QComboBox, QFileDialog, QLabel,
                             QMessageBox)
from PyQt6.QtCore import QThread, pyqtSignal, pyqtSlot, QSettings, Qt
from PyQt6.QtGui import QTextCharFormat, QColor, QIcon

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class SerialReaderThread(QThread):
    data_received = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self,connection):
        super().__init__()
        self.connection = connection
        self.running = False
        self.buffer = ""
        self.last_data_time = - float('inf')
        self.timeout = 0.05

    def run(self):
        try:
            self.running = True
            while self.running:
                if self.connection and self.connection.in_waiting:
                    byte_data = self.connection.read(self.connection.in_waiting)
                    self.handle_incoming_data(byte_data)

                if self.buffer and (time.time() - self.last_data_time) >= self.timeout:
                    self.flush_buffer()

                time.sleep(0.01)
        except serial.SerialException as e:
            self.error_occurred.emit(f"Error reading serial port:\n\n{e}\n")

    def handle_incoming_data(self, byte_data):
        decoded_data = self.decode_serial_data(byte_data)
        self.buffer += decoded_data

        if self.buffer.endswith('\n'):
            self.flush_buffer()
        else:
            self.last_data_time = time.time()

    def flush_buffer(self):
        if self.buffer:
            self.data_received.emit(self.buffer)
            self.buffer = ""

    def decode_serial_data(self, data):
        output = []
        for byte in data:
            if 32 <= byte <= 126 or byte in [9,10,13]:
                output.append(chr(byte))
            else:
                output.append(f"<{byte}>")
        return ''.join(output)

    def stop(self):
        self.running = False

class ScriptExecutionThread(QThread):
    error_occurred = pyqtSignal(str)
    log_request = pyqtSignal(str,str)

    def __init__(self, connection, script_path):
        super().__init__()
        self.connection = connection
        try:
            with open(script_path, "r") as file:
                self.script_content = file.read()
        except Exception as e:
            self.script_content = ''
            self.error_occurred.emit(f"Error reading script file:\n\n{e}\n")

    def run(self):
        if self.script_content and self.connection and self.connection.is_open:
            try:
                local_env = {'SERIAL': self.com_send, "print": self.log}
                exec(self.script_content, local_env)
            except Exception as e:
                self.error_occurred.emit(f"Error executing script file:\n\n{e}\n")
        else:
            self.error_occurred.emit("Serial port not open.")

    def com_send(self, data):
        if self.connection and self.connection.is_open:
            self.connection.write(data.encode())

    def log(self,*args, sep=' ', end='\n', color="yellow"):
        self.log_request.emit(sep.join(map(str, args)) + end,color)

class SerialMonitorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.serial_conn = None
        self.receive_thread = None
        self.selected_file = ''
        self.command_history = []
        self.history_index = 0
        self.command_history_changes = {}
        self.init_ui()
        self.refresh_com_ports()

    def init_ui(self):

        self.setWindowTitle("TermiCOM")
        self.setWindowIcon(QIcon(resource_path("assets\\icon.png")))

        layout = QVBoxLayout()

        config_layout = QHBoxLayout()
        self.com_port_combo = QComboBox()
        self.baud_rate_combo = QComboBox()
        self.baud_rate_combo.addItems(["9600", "19200", "38400", "57600", "115200"])
        self.refresh_button = QPushButton("Refresh")
        config_layout.addWidget(self.refresh_button)
        config_layout.addWidget(QLabel("COM Port:"))
        config_layout.addWidget(self.com_port_combo)
        config_layout.addWidget(QLabel("Baud Rate:"))
        config_layout.addWidget(self.baud_rate_combo)
        config_layout.addStretch()
        self.connect_button = QPushButton("Connect")
        self.connect_button.clicked.connect(self.toggle_serial_connection)
        self.connection_status_indicator = QLabel("  Disconnected  ")
        self.connection_status_indicator.setStyleSheet("background-color: red; width: 20px; height: 20px;")
        config_layout.addWidget(self.connect_button)
        config_layout.addWidget(self.connection_status_indicator)

        
        layout.addLayout(config_layout)

        self.textbox = QTextEdit()
        self.textbox.setReadOnly(True)
        layout.addWidget(self.textbox)

        command_layout = QHBoxLayout()
        self.command_input = QLineEdit()
        self.command_input.setPlaceholderText("Type command here...")
        self.command_input.returnPressed.connect(self.send_manual_command)
        self.command_input.installEventFilter(self)
        self.send_button = QPushButton("Send")
        self.send_button.clicked.connect(self.send_manual_command)
        command_layout.addWidget(self.command_input)
        command_layout.addWidget(self.send_button)
        layout.addLayout(command_layout)

        script_layout = QHBoxLayout()
        self.script_label = QLabel("No script selected.")
        self.script_button = QPushButton("Select Script")
        self.script_button.clicked.connect(self.select_script)
        self.run_button = QPushButton("Run Script")
        self.run_button.clicked.connect(self.run_selected_script)
        script_layout.addWidget(self.script_label)
        script_layout.addStretch()
        script_layout.addWidget(self.script_button)
        script_layout.addWidget(self.run_button)
        layout.addLayout(script_layout)

        self.setLayout(layout)

    def refresh_com_ports(self):
        self.com_port_combo.clear()
        ports = serial.tools.list_ports.comports()
        self.com_port_combo.addItems([port.device for port in ports])

    def toggle_serial_connection(self):
        if self.serial_conn and self.serial_conn.is_open:

            if self.receive_thread:
                self.receive_thread.stop()
                self.receive_thread.wait()
                self.receive_thread = None
            self.serial_conn.close()
            self.connection_status_indicator.setStyleSheet("background-color: red; width: 20px; height: 20px;")
            self.connection_status_indicator.setText("  Disconnected  ")
            self.connect_button.setText("Connect")

            self.com_port_combo.setEnabled(True)
            self.baud_rate_combo.setEnabled(True)

        else:

            port = self.com_port_combo.currentText()
            baud_rate = self.baud_rate_combo.currentText()

            try:
                self.serial_conn = serial.Serial(port, int(baud_rate), timeout=1)

                if self.serial_conn.is_open:

                    self.connection_status_indicator.setStyleSheet("background-color: green; width: 20px; height: 20px;")
                    self.connection_status_indicator.setText("  Connected  ")
                    self.connect_button.setText("Diconnect")

                    self.com_port_combo.setEnabled(False)
                    self.baud_rate_combo.setEnabled(False)

                    self.receive_thread = SerialReaderThread(self.serial_conn)
                    self.receive_thread.data_received.connect(self.update_textbox)
                    self.receive_thread.error_occurred.connect(self.show_error_popup)
                    self.receive_thread.start()

                else:
                    self.show_error_popup(f"Error opening serial port: {e}")

            except serial.SerialException as e:
                self.connection_status_indicator.setStyleSheet("background-color: red; width: 20px; height: 20px;")
                self.connection_status_indicator.setText("  Disconnected  ")
                self.connect_button.setText("Connect")

                self.com_port_combo.setEnabled(True)
                self.baud_rate_combo.setEnabled(True)

                self.show_error_popup(f"Error opening serial port:\n\n{e}\n")

    @pyqtSlot(str)
    def update_textbox(self, data):
        self.textbox.insertPlainText(data)

    @pyqtSlot(str,str)
    def log_textbox(self, data, color):
        cursor = self.textbox.textCursor()

        formt = QTextCharFormat()

        formt.setForeground(QColor(color))
        cursor.setCharFormat(formt)

        cursor.insertText(data)

        formt.setForeground(QColor("white"))
        cursor.setCharFormat(formt)

        self.textbox.setTextCursor(cursor)

    def show_error_popup(self, message):
        error_dialog = QMessageBox(self)
        error_dialog.setIcon(QMessageBox.Icon.Warning)
        error_dialog.setWindowTitle("Error")
        error_dialog.setText(message)
        error_dialog.exec()

    def select_script(self):
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        if file_dialog.exec():
            selected_file = file_dialog.selectedFiles()[0]
            if selected_file:
                self.script_label.setText(selected_file)
                self.selected_file = selected_file

    def run_selected_script(self):
        if self.selected_file:
            self.textbox.clear()
            self.script_thread = ScriptExecutionThread(self.serial_conn, self.selected_file)
            self.script_thread.error_occurred.connect(self.show_error_popup)
            self.script_thread.log_request.connect(self.log_textbox)
            self.script_thread.start()
        else:
            self.show_error_popup("No script loaded.\nPlease select a script first.")


    def send_manual_command(self):
        command = self.command_input.text()
        self.command_input.clear()
        if command:
            if (self.serial_conn and self.serial_conn.is_open) or True:
                # self.textbox.clear()
                # self.serial_conn.write(command.encode())

                if not self.command_history or self.command_history[-1] != command:
                    self.command_history.append(command)
                self.history_index = len(self.command_history)
                self.command_history_changes.clear()
                self.command_history_changes[self.history_index] = ''
            else:
                self.show_error_popup("Serial port not open.")

    def eventFilter(self, obj, event):
        if obj == self.command_input and event.type() == event.Type.KeyPress:

            if event.key() == Qt.Key.Key_Up:
                if self.history_index > 0:
                    self.command_history_changes[self.history_index] = self.command_input.text()
                    self.history_index -= 1
                    if self.history_index in self.command_history_changes:
                        self.command_input.setText(self.command_history_changes[self.history_index])
                    else:
                        self.command_input.setText(self.command_history[self.history_index])
                return True
            
            if event.key() == Qt.Key.Key_Down:
                if self.history_index < len(self.command_history):
                    self.command_history_changes[self.history_index] = self.command_input.text()
                    self.history_index += 1
                    self.command_input.setText(self.command_history_changes[self.history_index])
                return True

            if event.key() == Qt.Key.Key_Escape:
                self.history_index = len(self.command_history)
                self.command_input.clear()
                self.command_history_changes.clear()
                self.command_history_changes[self.history_index] = ''
                return True


        return super().eventFilter(obj, event)

    def closeEvent(self, event):
        if self.receive_thread:
            self.receive_thread.stop()
            self.receive_thread.wait()
            self.receive_thread = None
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
        event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = SerialMonitorApp()
    window.show()
    sys.exit(app.exec())
