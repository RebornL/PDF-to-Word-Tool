"""
图形用户界面模块
使用PyQt5实现直观的图形界面
"""

import sys
import os
from typing import Optional, List
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QProgressBar,
    QTextEdit, QGroupBox, QCheckBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QSplitter, QFrame, QStatusBar,
    QTabWidget, QPlainTextEdit, QSpinBox, QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QBrush, QTextCharFormat, QTextCursor

from pdf_converter import PDFConverter
from search_replace import SearchReplaceEngine, MatchResult, ReplacementPreview


class ConvertWorker(QThread):
    """PDF转换工作线程"""
    progress = pyqtSignal(int, int)  # 当前页, 总页数
    finished = pyqtSignal(bool, str)  # 成功, 消息

    def __init__(self, converter: PDFConverter, pdf_path: str, output_path: str):
        super().__init__()
        self.converter = converter
        self.pdf_path = pdf_path
        self.output_path = output_path

    def run(self):
        try:
            def progress_callback(current, total):
                self.progress.emit(current, total)

            success = self.converter.convert(
                self.pdf_path,
                self.output_path,
                progress_callback
            )
            self.finished.emit(success, "转换完成" if success else "转换已取消")
        except Exception as e:
            self.finished.emit(False, str(e))


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()

        # 初始化组件
        self.pdf_converter = PDFConverter()
        self.search_engine = SearchReplaceEngine()
        self.convert_worker: Optional[ConvertWorker] = None

        # 当前文件路径
        self.current_pdf_path = ""
        self.current_docx_path = ""
        self.current_output_path = ""

        # 搜索结果
        self.current_matches: List[MatchResult] = []
        self.current_previews: List[ReplacementPreview] = []

        # 初始化UI
        self.init_ui()

    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("PDF转Word工具 - 敏感词替换")
        self.setMinimumSize(1000, 700)
        self.resize(1200, 800)

        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 文件选择区域
        file_group = self.create_file_selection_group()
        main_layout.addWidget(file_group)

        # 创建分割器
        splitter = QSplitter(Qt.Vertical)

        # 上半部分：搜索替换区域
        search_widget = self.create_search_replace_widget()
        splitter.addWidget(search_widget)

        # 下半部分：预览区域
        preview_widget = self.create_preview_widget()
        splitter.addWidget(preview_widget)

        # 设置分割比例
        splitter.setSizes([400, 300])

        main_layout.addWidget(splitter, 1)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

    def create_file_selection_group(self) -> QGroupBox:
        """创建文件选择区域"""
        group = QGroupBox("文件选择")
        layout = QVBoxLayout(group)

        # PDF文件选择
        pdf_layout = QHBoxLayout()
        pdf_label = QLabel("PDF文件:")
        pdf_label.setMinimumWidth(80)
        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setPlaceholderText("选择要转换的PDF文件...")
        self.pdf_path_edit.setReadOnly(True)
        pdf_browse_btn = QPushButton("浏览...")
        pdf_browse_btn.clicked.connect(self.browse_pdf_file)
        pdf_browse_btn.setFixedWidth(80)

        pdf_layout.addWidget(pdf_label)
        pdf_layout.addWidget(self.pdf_path_edit, 1)
        pdf_layout.addWidget(pdf_browse_btn)
        layout.addLayout(pdf_layout)

        # 输出目录选择
        output_layout = QHBoxLayout()
        output_label = QLabel("输出目录:")
        output_label.setMinimumWidth(80)
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("选择输出目录（默认与PDF同目录）...")
        self.output_path_edit.setReadOnly(True)
        output_browse_btn = QPushButton("浏览...")
        output_browse_btn.clicked.connect(self.browse_output_dir)
        output_browse_btn.setFixedWidth(80)

        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_path_edit, 1)
        output_layout.addWidget(output_browse_btn)
        layout.addLayout(output_layout)

        # 转换按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.convert_btn = QPushButton("转换PDF为Word")
        self.convert_btn.setFixedWidth(150)
        self.convert_btn.clicked.connect(self.convert_pdf)
        btn_layout.addWidget(self.convert_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        return group

    def create_search_replace_widget(self) -> QWidget:
        """创建搜索替换区域"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        # 搜索设置组
        search_group = QGroupBox("搜索与替换设置")
        search_layout = QVBoxLayout(search_group)

        # 搜索词和替换词
        input_layout = QHBoxLayout()

        # 搜索词
        search_input_layout = QVBoxLayout()
        search_label = QLabel("搜索词:")
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("输入要搜索的关键词...")
        self.search_edit.setMinimumWidth(200)
        self.search_edit.returnPressed.connect(self.search_keyword)
        search_input_layout.addWidget(search_label)
        search_input_layout.addWidget(self.search_edit)

        # 替换词
        replace_input_layout = QVBoxLayout()
        replace_label = QLabel("替换为:")
        self.replace_edit = QLineEdit()
        self.replace_edit.setPlaceholderText("输入替换后的文本...")
        self.replace_edit.setMinimumWidth(200)
        replace_input_layout.addWidget(replace_label)
        replace_input_layout.addWidget(self.replace_edit)

        input_layout.addLayout(search_input_layout)
        input_layout.addLayout(replace_input_layout)

        # 匹配选项
        options_layout = QHBoxLayout()
        self.case_sensitive_cb = QCheckBox("区分大小写")
        self.whole_word_cb = QCheckBox("全词匹配")
        options_layout.addWidget(self.case_sensitive_cb)
        options_layout.addWidget(self.whole_word_cb)
        options_layout.addStretch()

        # 按钮
        self.search_btn = QPushButton("搜索")
        self.search_btn.setFixedWidth(80)
        self.search_btn.clicked.connect(self.search_keyword)
        options_layout.addWidget(self.search_btn)

        self.preview_btn = QPushButton("预览替换")
        self.preview_btn.setFixedWidth(80)
        self.preview_btn.clicked.connect(self.preview_replacements)
        self.preview_btn.setEnabled(False)
        options_layout.addWidget(self.preview_btn)

        input_layout.addLayout(options_layout)
        search_layout.addLayout(input_layout)

        # 批量替换列表
        batch_label = QLabel("批量替换列表（每行一个：搜索词=替换词）:")
        search_layout.addWidget(batch_label)

        self.batch_edit = QPlainTextEdit()
        self.batch_edit.setPlaceholderText("例如:\n张三=李四\n电话=联系方式\n身份证=证件号码")
        self.batch_edit.setMaximumHeight(80)
        search_layout.addWidget(self.batch_edit)

        layout.addWidget(search_group)

        # 操作按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.apply_selected_btn = QPushButton("替换选中项")
        self.apply_selected_btn.setFixedWidth(100)
        self.apply_selected_btn.clicked.connect(self.replace_selected)
        self.apply_selected_btn.setEnabled(False)
        btn_layout.addWidget(self.apply_selected_btn)

        self.apply_all_btn = QPushButton("替换全部")
        self.apply_all_btn.setFixedWidth(100)
        self.apply_all_btn.clicked.connect(self.replace_all)
        self.apply_all_btn.setEnabled(False)
        btn_layout.addWidget(self.apply_all_btn)

        self.batch_replace_btn = QPushButton("批量替换")
        self.batch_replace_btn.setFixedWidth(100)
        self.batch_replace_btn.clicked.connect(self.batch_replace)
        self.batch_replace_btn.setEnabled(False)
        btn_layout.addWidget(self.batch_replace_btn)

        self.save_btn = QPushButton("保存文档")
        self.save_btn.setFixedWidth(100)
        self.save_btn.clicked.connect(self.save_document)
        self.save_btn.setEnabled(False)
        btn_layout.addWidget(self.save_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        return widget

    def create_preview_widget(self) -> QWidget:
        """创建预览区域"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        # 预览组
        preview_group = QGroupBox("搜索结果与替换预览")
        preview_layout = QVBoxLayout(preview_group)

        # 结果统计
        self.result_label = QLabel("共找到 0 处匹配")
        preview_layout.addWidget(self.result_label)

        # 结果表格
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(6)
        self.result_table.setHorizontalHeaderLabels([
            "选择", "位置", "匹配文本", "替换为", "上下文", "预览"
        ])

        # 设置表头
        header = self.result_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Interactive)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.Interactive)

        self.result_table.setColumnWidth(0, 50)
        self.result_table.setColumnWidth(1, 120)
        self.result_table.setColumnWidth(2, 150)
        self.result_table.setColumnWidth(3, 150)
        self.result_table.setColumnWidth(5, 200)

        preview_layout.addWidget(self.result_table)

        # 全选/取消全选按钮
        select_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.setFixedWidth(60)
        self.select_all_btn.clicked.connect(self.select_all_results)
        select_layout.addWidget(self.select_all_btn)

        self.deselect_all_btn = QPushButton("取消全选")
        self.deselect_all_btn.setFixedWidth(70)
        self.deselect_all_btn.clicked.connect(self.deselect_all_results)
        select_layout.addWidget(self.deselect_all_btn)

        select_layout.addStretch()
        preview_layout.addLayout(select_layout)

        layout.addWidget(preview_group)

        return widget

    def browse_pdf_file(self):
        """浏览选择PDF文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择PDF文件",
            "",
            "PDF文件 (*.pdf);;所有文件 (*.*)"
        )

        if file_path:
            self.current_pdf_path = file_path
            self.pdf_path_edit.setText(file_path)

            # 自动设置输出目录
            if not self.output_path_edit.text():
                output_dir = os.path.dirname(file_path)
                self.output_path_edit.setText(output_dir)

            self.status_bar.showMessage(f"已选择PDF文件: {os.path.basename(file_path)}")

    def browse_output_dir(self):
        """浏览选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "选择输出目录",
            ""
        )

        if dir_path:
            self.output_path_edit.setText(dir_path)
            self.status_bar.showMessage(f"输出目录: {dir_path}")

    def convert_pdf(self):
        """转换PDF为Word"""
        if not self.current_pdf_path:
            QMessageBox.warning(self, "警告", "请先选择PDF文件！")
            return

        if not os.path.exists(self.current_pdf_path):
            QMessageBox.warning(self, "警告", "PDF文件不存在！")
            return

        # 确定输出路径
        output_dir = self.output_path_edit.text()
        if not output_dir:
            output_dir = os.path.dirname(self.current_pdf_path)

        pdf_name = os.path.splitext(os.path.basename(self.current_pdf_path))[0]
        self.current_docx_path = os.path.join(output_dir, f"{pdf_name}.docx")
        self.current_output_path = self.current_docx_path

        # 禁用按钮
        self.convert_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_bar.showMessage("正在转换PDF...")

        # 启动转换线程
        self.convert_worker = ConvertWorker(
            self.pdf_converter,
            self.current_pdf_path,
            self.current_docx_path
        )
        self.convert_worker.progress.connect(self.on_convert_progress)
        self.convert_worker.finished.connect(self.on_convert_finished)
        self.convert_worker.start()

    def on_convert_progress(self, current: int, total: int):
        """转换进度回调"""
        if total > 0:
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(current)
            self.status_bar.showMessage(f"正在转换: {current}/{total} 页")

    def on_convert_finished(self, success: bool, message: str):
        """转换完成回调"""
        self.progress_bar.setVisible(False)
        self.convert_btn.setEnabled(True)

        if success:
            self.status_bar.showMessage(f"转换完成: {self.current_docx_path}")

            # 加载文档进行搜索替换
            try:
                self.search_engine.load_document(self.current_docx_path)
                self.preview_btn.setEnabled(True)
                self.save_btn.setEnabled(True)
                self.batch_replace_btn.setEnabled(True)

                # 获取文档统计
                stats = self.search_engine.get_document_statistics()
                self.result_label.setText(
                    f"文档已加载 - 段落数: {stats.get('paragraphs', 0)}, "
                    f"表格数: {stats.get('tables', 0)}"
                )

                QMessageBox.information(
                    self,
                    "转换成功",
                    f"PDF已成功转换为Word文档！\n\n输出文件: {self.current_docx_path}"
                )
            except Exception as e:
                QMessageBox.warning(self, "警告", f"加载文档失败: {str(e)}")
        else:
            self.status_bar.showMessage(f"转换失败: {message}")
            QMessageBox.warning(self, "转换失败", message)

    def search_keyword(self):
        """搜索关键词"""
        keyword = self.search_edit.text().strip()
        if not keyword:
            QMessageBox.warning(self, "警告", "请输入搜索关键词！")
            return

        if not self.search_engine.document:
            QMessageBox.warning(self, "警告", "请先转换PDF文件！")
            return

        try:
            self.current_matches = self.search_engine.search(
                keyword,
                case_sensitive=self.case_sensitive_cb.isChecked(),
                whole_word=self.whole_word_cb.isChecked()
            )

            self.update_result_table()
            self.preview_btn.setEnabled(len(self.current_matches) > 0)
            self.apply_all_btn.setEnabled(len(self.current_matches) > 0)

            self.result_label.setText(f"共找到 {len(self.current_matches)} 处匹配")
            self.status_bar.showMessage(f"搜索完成，找到 {len(self.current_matches)} 处匹配")

        except Exception as e:
            QMessageBox.warning(self, "搜索失败", str(e))

    def preview_replacements(self):
        """预览替换结果"""
        keyword = self.search_edit.text().strip()
        replacement = self.replace_edit.text()

        if not keyword:
            QMessageBox.warning(self, "警告", "请输入搜索关键词！")
            return

        if not self.search_engine.document:
            QMessageBox.warning(self, "警告", "请先转换PDF文件！")
            return

        try:
            self.current_previews = self.search_engine.preview_replacements(
                keyword,
                replacement,
                case_sensitive=self.case_sensitive_cb.isChecked(),
                whole_word=self.whole_word_cb.isChecked()
            )

            self.update_preview_table()
            self.apply_selected_btn.setEnabled(len(self.current_previews) > 0)
            self.apply_all_btn.setEnabled(len(self.current_previews) > 0)

            self.result_label.setText(f"共 {len(self.current_previews)} 处可替换")
            self.status_bar.showMessage("预览完成")

        except Exception as e:
            QMessageBox.warning(self, "预览失败", str(e))

    def update_result_table(self):
        """更新搜索结果表格"""
        self.result_table.setRowCount(len(self.current_matches))

        for i, match in enumerate(self.current_matches):
            # 选择框
            check_item = QTableWidgetItem()
            check_item.setCheckState(Qt.Checked)
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            self.result_table.setItem(i, 0, check_item)

            # 位置
            self.result_table.setItem(i, 1, QTableWidgetItem(match.location))

            # 匹配文本
            self.result_table.setItem(i, 2, QTableWidgetItem(match.match_text))

            # 替换为（暂时为空）
            self.result_table.setItem(i, 3, QTableWidgetItem(""))

            # 上下文
            context_item = QTableWidgetItem(match.context)
            self.result_table.setItem(i, 4, context_item)

            # 预览（暂时为空）
            self.result_table.setItem(i, 5, QTableWidgetItem(""))

    def update_preview_table(self):
        """更新预览表格"""
        self.result_table.setRowCount(len(self.current_previews))

        for i, preview in enumerate(self.current_previews):
            # 选择框
            check_item = QTableWidgetItem()
            check_item.setCheckState(Qt.Checked)
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            self.result_table.setItem(i, 0, check_item)

            # 位置
            self.result_table.setItem(i, 1, QTableWidgetItem(preview.match.location))

            # 匹配文本
            match_item = QTableWidgetItem(preview.match.match_text)
            match_item.setBackground(QBrush(QColor(255, 255, 200)))  # 高亮显示
            self.result_table.setItem(i, 2, match_item)

            # 替换为
            replace_item = QTableWidgetItem(preview.replacement)
            replace_item.setBackground(QBrush(QColor(200, 255, 200)))  # 高亮显示
            self.result_table.setItem(i, 3, replace_item)

            # 上下文
            self.result_table.setItem(i, 4, QTableWidgetItem(preview.match.context))

            # 预览
            preview_text = f"...{preview.after[max(0, preview.match.start_pos-20):preview.match.start_pos + len(preview.replacement) + 20]}..."
            self.result_table.setItem(i, 5, QTableWidgetItem(preview_text))

    def select_all_results(self):
        """全选结果"""
        for i in range(self.result_table.rowCount()):
            item = self.result_table.item(i, 0)
            if item:
                item.setCheckState(Qt.Checked)

    def deselect_all_results(self):
        """取消全选"""
        for i in range(self.result_table.rowCount()):
            item = self.result_table.item(i, 0)
            if item:
                item.setCheckState(Qt.Unchecked)

    def get_selected_indices(self) -> List[int]:
        """获取选中的索引"""
        indices = []
        for i in range(self.result_table.rowCount()):
            item = self.result_table.item(i, 0)
            if item and item.checkState() == Qt.Checked:
                indices.append(i)
        return indices

    def replace_selected(self):
        """替换选中的内容"""
        if not self.current_previews:
            return

        selected_indices = self.get_selected_indices()
        if not selected_indices:
            QMessageBox.warning(self, "警告", "请选择要替换的内容！")
            return

        keyword = self.search_edit.text().strip()
        replacement = self.replace_edit.text()

        reply = QMessageBox.question(
            self,
            "确认替换",
            f"确定要替换选中的 {len(selected_indices)} 处内容吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.do_replace(keyword, replacement, selected_indices)

    def replace_all(self):
        """替换全部"""
        if not self.current_previews:
            return

        keyword = self.search_edit.text().strip()
        replacement = self.replace_edit.text()

        reply = QMessageBox.question(
            self,
            "确认替换",
            f"确定要替换全部 {len(self.current_previews)} 处内容吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.do_replace(keyword, replacement, None)

    def do_replace(self, keyword: str, replacement: str, selected_indices: Optional[List[int]]):
        """执行替换操作"""
        try:
            count = self.search_engine.replace(
                keyword,
                replacement,
                case_sensitive=self.case_sensitive_cb.isChecked(),
                whole_word=self.whole_word_cb.isChecked(),
                selected_indices=selected_indices
            )

            self.status_bar.showMessage(f"已替换 {count} 处")

            # 重新搜索更新结果
            self.search_keyword()

            QMessageBox.information(
                self,
                "替换完成",
                f"成功替换了 {count} 处内容！\n\n请记得保存文档。"
            )

        except Exception as e:
            QMessageBox.warning(self, "替换失败", str(e))

    def batch_replace(self):
        """批量替换"""
        batch_text = self.batch_edit.toPlainText().strip()
        if not batch_text:
            QMessageBox.warning(self, "警告", "请输入批量替换列表！")
            return

        if not self.search_engine.document:
            QMessageBox.warning(self, "警告", "请先转换PDF文件！")
            return

        # 解析批量替换列表
        replace_pairs = []
        for line in batch_text.split('\n'):
            line = line.strip()
            if '=' in line:
                parts = line.split('=', 1)
                if len(parts) == 2:
                    search_word = parts[0].strip()
                    replace_word = parts[1].strip()
                    if search_word:
                        replace_pairs.append((search_word, replace_word))

        if not replace_pairs:
            QMessageBox.warning(self, "警告", "未找到有效的替换规则！")
            return

        # 显示预览
        preview_text = "以下替换将被执行:\n\n"
        total_matches = 0

        for search_word, replace_word in replace_pairs:
            matches = self.search_engine.search(search_word)
            total_matches += len(matches)
            preview_text += f"'{search_word}' -> '{replace_word}': {len(matches)} 处\n"

        preview_text += f"\n总计: {total_matches} 处将被替换"

        reply = QMessageBox.question(
            self,
            "确认批量替换",
            preview_text + "\n\n确定继续吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            total_count = 0
            for search_word, replace_word in replace_pairs:
                count = self.search_engine.replace(search_word, replace_word)
                total_count += count

            self.status_bar.showMessage(f"批量替换完成，共替换 {total_count} 处")

            QMessageBox.information(
                self,
                "批量替换完成",
                f"成功替换了 {total_count} 处内容！\n\n请记得保存文档。"
            )

            # 清空批量替换列表
            self.batch_edit.clear()

    def save_document(self):
        """保存文档"""
        if not self.search_engine.document:
            QMessageBox.warning(self, "警告", "没有可保存的文档！")
            return

        # 选择保存路径
        default_name = self.current_output_path
        if not default_name:
            default_name = "output.docx"

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存Word文档",
            default_name,
            "Word文档 (*.docx);;所有文件 (*.*)"
        )

        if file_path:
            try:
                self.search_engine.save_document(file_path)
                self.status_bar.showMessage(f"文档已保存: {file_path}")

                QMessageBox.information(
                    self,
                    "保存成功",
                    f"文档已成功保存！\n\n保存位置: {file_path}"
                )
            except Exception as e:
                QMessageBox.warning(self, "保存失败", str(e))

    def closeEvent(self, event):
        """窗口关闭事件"""
        if self.convert_worker and self.convert_worker.isRunning():
            reply = QMessageBox.question(
                self,
                "确认退出",
                "PDF转换正在进行中，确定要退出吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if reply == QMessageBox.No:
                event.ignore()
                return

            self.pdf_converter.cancel()
            self.convert_worker.wait()

        event.accept()


def main():
    """主函数"""
    app = QApplication(sys.argv)

    # 设置应用样式
    app.setStyle('Fusion')

    # 创建并显示主窗口
    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
