# -*- coding: utf-8 -*-
import datetime
import os
import re
import sys
import traceback
from html.parser import HTMLParser

import PySide6
import xlsxwriter
from PySide6.QtCore import QObject, QThread, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMainWindow

from log_utils import logconfig
from mainwindow import Ui_MainWindow


logger = logconfig.get_logger()

supported_extensions = ['.html']


class MyHTMLParser(HTMLParser):
    """
    Ref: https://stackoverflow.com/questions/8477627/iteratively-parsing-html-with-lxml/8484265#8484265
    """

    def __init__(self, callback):
        self.finished = False
        self.in_table = False
        self.in_row = False
        self.in_cell = False
        self.current_row = []
        self.current_cell = None
        self.row_idx = 0
        self.callback = callback
        HTMLParser.__init__(self)

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if not self.in_table:
            if tag == 'table':
                self.in_table = True
        else:
            if tag == 'tr':
                self.in_row = True
            # FIXME: 'th'表头与'td'分开处理
            elif tag == 'td' or tag == 'th':
                self.in_cell = True

    def handle_endtag(self, tag):
        if tag == 'tr':
            if self.in_table:
                if self.in_row:
                    self.in_row = False
                    self.callback(self.row_idx, self.current_row)
                    self.current_row = []
                    self.row_idx += 1
        elif tag == 'td' or tag == 'th':
            if self.in_table:
                if self.in_cell:
                    self.in_cell = False
                    self.current_row.append(self.current_cell)
                    self.current_cell = None

        elif (tag == 'table') and self.in_table:
            self.finished = True

    def handle_data(self, data):
        if self.in_cell:
            self.current_cell = data.strip() if data else data


class MainWindow(QMainWindow):
    drop = Signal(str)

    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowIcon(QIcon('favicon.png'))
        self.setWindowTitle('html2xlsx')
        self.drop.connect(self.open_html_file)
        self.ui.cancelButton.clicked.connect(self.on_canceled)
        self.ui.openOutputButton.clicked.connect(self.on_open_output)
        self.setAcceptDrops(True)
        self.ui.progressBar.setValue(0)
        self.progress_max = 0

    def set_progress_max(self, max_value: int):
        self.progress_max = max_value
        self.ui.progressBar.setMaximum(max_value)

    def log_msg(self, msg: str) -> None:
        logger.info(msg)
        self.ui.plainTextEdit.appendPlainText(
            f'{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}  {msg}')

    def open_html_file(self, html_path):
        self.log_msg(f'open_html_file {html_path}')
        xlsx_path = os.path.splitext(html_path)[0] + '.xlsx'
        self.html_path = html_path
        self.xlsx_path = xlsx_path

        self.ui.plainTextEdit.clear()
        self.ui.statusbar.showMessage(f'正在处理 {html_path}')
        self.ui.outputEdit.setText('')

        self.ui.progressBar.setValue(0)

        def report_progress(progress: int):
            self.log_msg(f'TR {progress}')
            self.ui.progressBar.setValue(progress)

        # Step 2: Create a QThread object
        self.thread = QThread()
        # Step 3: Create a worker object
        self.worker = ParserWorker(html_path, xlsx_path)
        # Step 4: Move worker to the thread
        self.worker.moveToThread(self.thread)
        # Step 5: Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(lambda x: self.thread.quit())
        # 这个会触发：RuntimeError: wrapped C/C++ object of type QThread has been deleted
        # self.worker.finished.connect(self.worker.deleteLater)
        # self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(report_progress)
        # Step 6: Start the thread
        self.thread.start()

        def thread_finished():
            self.setAcceptDrops(True)
            self.log_msg('thread finished')

        def worker_finished(success: bool):
            if success:
                self.ui.progressBar.setValue(self.progress_max)
                self.log_msg('worker finished')
                self.ui.statusbar.showMessage(f'处理完成: {self.xlsx_path}')
                self.ui.outputEdit.setText(self.xlsx_path)
            else:
                self.ui.progressBar.setValue(0)
                self.log_msg('worker suspended.')
                self.ui.statusbar.showMessage(f'处理中止: {self.html_path}')
                self.ui.outputEdit.setText('')

        # Final resets
        self.setAcceptDrops(False)
        self.thread.finished.connect(thread_finished)
        self.worker.finished.connect(worker_finished)
        self.worker.msg.connect(self.log_msg)
        self.worker.rowcount.connect(self.set_progress_max)

    def on_canceled(self):
        self.log_msg('html parsing canceled')
        if self.thread and self.thread.isRunning():
            self.thread.requestInterruption()
            self.thread.quit()
            self.thread.wait()
            self.thread = None
            self.worker = None
            self.setAcceptDrops(True)
            self.ui.statusbar.showMessage(f'已取消 {self.html_path}')
            self.ui.progressBar.setValue(0)

    def on_open_output(self):
        xlsx_path = self.ui.outputEdit.text()
        if os.path.exists(xlsx_path):
            os.startfile(xlsx_path)

    def dragEnterEvent(self, event: PySide6.QtGui.QDragEnterEvent) -> None:
        self.log_msg('drag-enter')
        if event.mimeData().hasUrls() and len(event.mimeData().urls()) == 1:
            # 只接收一个文件/次
            event.accept()
            return

        event.ignore()

    def dropEvent(self, event: PySide6.QtGui.QDropEvent) -> None:
        self.log_msg('drop')
        if event.mimeData().hasUrls():
            file_path = event.mimeData().urls()[0].toLocalFile()
            if os.path.splitext(file_path)[1] in supported_extensions:
                self.drop.emit(file_path)
                event.accept()
                return

        event.ignore()


class MyInterrupt(Exception):
    pass


class ParserWorker(QObject):
    finished = Signal(bool)
    progress = Signal(int)
    msg = Signal(str)
    rowcount = Signal(int)

    def __init__(self, html_path, xlsx_path, parent=None):
        super().__init__(parent)
        self.html_path: str = html_path
        self.xlsx_path: str = xlsx_path

        if os.path.exists(xlsx_path):
            logger.info(f'remove {xlsx_path}')
            os.remove(xlsx_path)

        # Create a workbook and add a worksheet.
        # https://xlsxwriter.readthedocs.io/working_with_memory.html
        self.workbook = xlsxwriter.Workbook(
            xlsx_path, {'constant_memory': True})
        self.worksheet = self.workbook.add_worksheet()

        self.parser = MyHTMLParser(self.write_row)

    def write_row(self, row, data):
        if QThread.currentThread().isInterruptionRequested():
            raise MyInterrupt('Interrupted')
        self.worksheet.write_row(row, 0, data)
        if row % 100 == 0:
            self.progress.emit(row)

    def run(self):
        encoding = 'utf-8'
        success = False
        try:
            with open(self.html_path, encoding='utf-8', errors='ignore') as f:
                data = f.read().lower()
                # 获取html文件的编码
                match = re.search(
                    r'meta.+charset=([\w-]+)', data)
                if match:
                    encoding = match.group(1)
                self.msg.emit(f'html encoding: {encoding}')
                # 计算表格行数
                row_count = data.count('<tr>')
                self.rowcount.emit(row_count)
                self.msg.emit(f'row count: {row_count}')

            self.parser.feed(open(self.html_path, encoding=encoding).read())
            success = True
        except MyInterrupt as e:
            self.msg.emit(f'worker interrupted.')
        except Exception as e:
            self.msg.emit(
                f'{"".join(traceback.format_exception(Exception, e, e.__traceback__))}')
        finally:
            self.parser.close()
            self.workbook.close()
            if not success:
                # 删除xlsx文件，要在`workbook.close()`之后
                if os.path.exists(self.xlsx_path):
                    os.remove(self.xlsx_path)
            self.finished.emit(success)


if __name__ == '__main__':
    logger.info(f'args: {sys.argv}')

    app = QApplication(sys.argv)
    with open("style.qss", "r", encoding='utf-8') as f:
        _style = f.read()
        app.setStyleSheet(_style)
    window = MainWindow()
    window.show()
    exit_code = app.exec()
    sys.exit(exit_code)
