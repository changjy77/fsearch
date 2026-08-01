#!/usr/bin/env python3
"""
fsearch GUI - PyQt5 기반 파일 검색 도구
"""

import sys
import os
import re
import time
import subprocess
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List
from collections import defaultdict

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QCheckBox,
    QSpinBox, QProgressBar, QComboBox, QTabWidget, QTableWidget, QTableWidgetItem,
    QSplitter, QMessageBox, QStyledItemDelegate
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QRect
from PyQt5.QtGui import QIcon, QColor, QFont, QCursor, QPainter

# 파일 형식별 텍스트 추출 라이브러리
try:
    from docx import Document
except ImportError:
    Document = None

try:
    from PyPDF2 import PdfReader
except ImportError:
    PdfReader = None

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

try:
    import html2text
except ImportError:
    html2text = None


class SearchResultDelegate(QStyledItemDelegate):
    """검색 결과에서 검색 단어를 굵게 표시하는 delegate"""

    def __init__(self, keyword, use_regex=False):
        super().__init__()
        self.keyword = keyword
        self.use_regex = use_regex
        if use_regex:
            try:
                self.pattern = re.compile(keyword, re.IGNORECASE)
            except:
                self.pattern = None
        else:
            self.pattern = None

    def paint(self, painter, option, index):
        """검색 단어를 굵게 표시하여 렌더링"""
        text = index.data(Qt.DisplayRole)
        if not text:
            super().paint(painter, option, index)
            return

        # 검색 단어 찾기
        if self.use_regex and self.pattern:
            matches = list(self.pattern.finditer(str(text)))
        else:
            matches = []
            keyword_lower = self.keyword.lower()
            text_lower = str(text).lower()
            start = 0
            while True:
                pos = text_lower.find(keyword_lower, start)
                if pos == -1:
                    break
                matches.append((pos, pos + len(self.keyword)))
                start = pos + 1

        if not matches:
            super().paint(painter, option, index)
            return

        # 배경색 설정
        painter.fillRect(option.rect, option.palette.base())

        # 텍스트와 bold 텍스트 렌더링
        painter.setPen(option.palette.text().color())

        bold_font = QFont(option.font)
        bold_font.setBold(True)

        x = option.rect.left() + 2
        y = option.rect.top() + option.fontMetrics.ascent() + 2

        text_str = str(text)
        last_pos = 0

        for start, end in matches:
            # 일반 텍스트
            if start > last_pos:
                normal_text = text_str[last_pos:start]
                painter.setFont(option.font)
                fm = painter.fontMetrics()
                painter.drawText(x, y, normal_text)
                x += fm.width(normal_text)

            # 굵은 텍스트
            bold_text = text_str[start:end]
            painter.setFont(bold_font)
            fm_bold = painter.fontMetrics()
            painter.drawText(x, y, bold_text)
            x += fm_bold.width(bold_text)

            last_pos = end

        # 남은 텍스트
        if last_pos < len(text_str):
            remaining_text = text_str[last_pos:]
            painter.setFont(option.font)
            fm = painter.fontMetrics()
            painter.drawText(x, y, remaining_text)


class SearchWorker(QThread):
    """검색을 별도 스레드에서 실행"""
    progress = pyqtSignal(int)
    result_found = pyqtSignal(dict)  # 결과를 실시간으로 전송
    finished = pyqtSignal(list)
    status = pyqtSignal(str)
    error = pyqtSignal(str)
    excluded_files_updated = pyqtSignal(list)  # 제외된 파일 목록 업데이트

    def __init__(self, keyword, path, ignore_dirs, name_only, content_only, use_regex, max_workers):
        super().__init__()
        self.keyword = keyword
        self.path = path
        self.ignore_dirs = set(ignore_dirs)
        self.name_only = name_only
        self.content_only = content_only
        self.use_regex = use_regex
        self.max_workers = max_workers
        self.results = []
        self.excluded_files = []  # 제외된 파일 목록

    def run(self):
        """검색 실행"""
        try:
            # 파일 수집
            self.status.emit("📂 파일을 수집 중입니다...")
            files = self._collect_files()

            if not files:
                self.status.emit("❌ 검색할 파일이 없습니다.")
                self.finished.emit([])
                return

            self.status.emit(f"✅ {len(files)}개 파일 발견. 검색 중...")

            # 정규식 컴파일
            if self.use_regex:
                try:
                    regex = re.compile(self.keyword, re.IGNORECASE)
                except re.error as e:
                    self.error.emit(f"정규식 오류: {e}")
                    return
            else:
                regex = None

            # 병렬 검색
            results = []
            processed = 0

            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                futures = {executor.submit(self._search_file, f, regex): f for f in files}

                for future in as_completed(futures):
                    try:
                        file_results = future.result()
                        results.extend(file_results)
                        # 실시간으로 각 결과 전송
                        for result in file_results:
                            self.result_found.emit(result)
                    except:
                        pass

                    processed += 1
                    self.progress.emit(int((processed / len(files)) * 100))

            self.finished.emit(results)
            self.status.emit(f"✅ 검색 완료: {len(results)}개 결과")

        except Exception as e:
            self.error.emit(f"오류 발생: {str(e)}")

    def _collect_files(self) -> List[Path]:
        """파일 수집 - 지정된 파일 형식만"""
        # 허용된 파일 확장자
        allowed_extensions = {
            '.doc', '.docx',      # 워드파일
            '.hwp',               # 한글파일
            '.pdf',               # PDF파일
            '.xls', '.xlsx',      # 엑셀파일
            '.txt',               # 텍스트파일
            '.html', '.htm',      # HTML파일
            '.md'                 # 마크다운파일
        }

        files = []
        self.excluded_files = []  # 초기화

        for root, dirs, filenames in os.walk(self.path):
            dirs[:] = [d for d in dirs if d not in self.ignore_dirs]
            for filename in filenames:
                file_path = Path(root) / filename
                # 허용된 확장자만 수집
                if file_path.suffix.lower() in allowed_extensions:
                    files.append(file_path)
                else:
                    # 제외된 파일 추적
                    self.excluded_files.append(str(file_path))

        # 제외된 파일 목록 업데이트 신호
        self.excluded_files_updated.emit(self.excluded_files)

        return files

    def _extract_text(self, file_path: Path) -> str:
        """파일 형식별로 텍스트 추출"""
        ext = file_path.suffix.lower()

        try:
            if ext == '.txt' or ext == '.md':
                # 텍스트/마크다운 파일
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()

            elif ext == '.html' or ext == '.htm':
                # HTML 파일
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()

            elif ext == '.docx' and Document:
                # Word 문서
                try:
                    doc = Document(file_path)
                    text = '\n'.join([para.text for para in doc.paragraphs])
                    return text
                except:
                    return ""

            elif ext == '.pdf' and PdfReader:
                # PDF 파일
                try:
                    reader = PdfReader(file_path)
                    text = ""
                    for page in reader.pages:
                        text += page.extract_text() + "\n"
                    return text
                except:
                    return ""

            elif ext in ['.xlsx', '.xls'] and load_workbook:
                # Excel 파일
                try:
                    if ext == '.xlsx':
                        wb = load_workbook(file_path, data_only=True)
                        text = ""
                        for ws in wb.sheetnames:
                            sheet = wb[ws]
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value:
                                        text += str(cell.value) + " "
                                text += "\n"
                        return text
                    else:
                        # .xls 파일은 openpyxl로 지원하지 않음
                        import xlrd
                        wb = xlrd.open_workbook(file_path)
                        text = ""
                        for sheet in wb.sheets():
                            for row in sheet.get_rows():
                                for cell in row:
                                    text += str(cell.value) + " "
                                text += "\n"
                        return text
                except:
                    return ""

            else:
                return ""

        except Exception as e:
            return ""

    def _search_file(self, file_path: Path, regex):
        """단일 파일 검색"""
        results = []

        # 파일 정보 수집
        try:
            stat = file_path.stat()
            file_size = stat.st_size
            mod_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat.st_mtime))
        except:
            file_size = 0
            mod_time = "Unknown"

        # 파일명과 폴더 경로 분리
        filename = file_path.name
        folder_path = str(file_path.parent)

        # 파일 아이콘 가져오기
        icon = SearchWorker.get_file_icon(filename)
        filename_with_icon = f"{icon} {filename}"

        total_match_count = 0
        filename_matched = False

        # 파일명 검색
        if not self.content_only:
            if self._match_keyword(file_path.name, regex):
                total_match_count += 1
                filename_matched = True

        # 파일 내용 검색 - 파일 형식별로 텍스트 추출
        content_match_count = 0
        if not self.name_only:
            # 파일에서 텍스트 추출
            text = self._extract_text(file_path)
            if text:
                # 추출된 텍스트를 줄 단위로 검색
                for line in text.split('\n'):
                    if self._match_keyword(line, regex):
                        total_match_count += 1
                        content_match_count += 1

        # 매칭이 있으면 파일당 하나의 결과만 추가
        if total_match_count > 0:
            results.append({
                'type': 'file_summary',
                'filename': filename_with_icon,
                'folder_path': folder_path,
                'full_path': str(file_path),
                'size': file_size,
                'modified': mod_time,
                'match_count': total_match_count
            })

        return results

    def _match_keyword(self, text: str, regex):
        """키워드 매칭"""
        if regex:
            return regex.search(text) is not None
        else:
            return self.keyword.lower() in text.lower()

    def _is_binary(self, path: Path) -> bool:
        """바이너리 파일 확인"""
        binary_exts = {'exe', 'dll', 'so', 'png', 'jpg', 'zip', 'db', 'pdf'}
        return path.suffix.lower().lstrip('.') in binary_exts

    @staticmethod
    def get_file_icon(filename: str) -> str:
        """파일 형식에 따라 아이콘 반환"""
        ext = Path(filename).suffix.lower()

        icon_map = {
            # 문서
            '.txt': '📄',
            '.md': '📝',
            '.pdf': '📕',
            '.doc': '📘',
            '.docx': '📘',
            '.xls': '📊',
            '.xlsx': '📊',
            '.csv': '📋',
            '.json': '⚙️',
            '.xml': '⚙️',
            '.yaml': '⚙️',
            '.yml': '⚙️',

            # 코드
            '.py': '🐍',
            '.js': '⚡',
            '.ts': '📘',
            '.tsx': '⚛️',
            '.jsx': '⚛️',
            '.cpp': '⚙️',
            '.c': '⚙️',
            '.h': '⚙️',
            '.java': '☕',
            '.cs': '💠',
            '.go': '🐹',
            '.rs': '🦀',
            '.rb': '💎',
            '.php': '🐘',
            '.html': '🌐',
            '.css': '🎨',
            '.scss': '🎨',
            '.sql': '🗄️',
            '.sh': '🔧',
            '.bat': '🔧',
            '.ps1': '🔧',

            # 이미지
            '.png': '🖼️',
            '.jpg': '🖼️',
            '.jpeg': '🖼️',
            '.gif': '🎬',
            '.svg': '🎨',
            '.ico': '🎯',

            # 압축
            '.zip': '📦',
            '.rar': '📦',
            '.7z': '📦',
            '.tar': '📦',
            '.gz': '📦',

            # 실행파일
            '.exe': '⚙️',
            '.dll': '⚙️',
            '.so': '⚙️',

            # 기타
            '.git': '🔀',
            '.env': '🔐',
            '.log': '📋',
        }

        return icon_map.get(ext, '📁')


class FSearchGUI(QMainWindow):
    """fsearch GUI 메인 윈도우"""

    def __init__(self):
        super().__init__()
        self.search_worker = None
        self.results = []
        self.excluded_files = []  # 제외된 파일 목록
        self.init_ui()

    def init_ui(self):
        """UI 초기화"""
        self.setWindowTitle("🔍 fsearch - 파일 검색 도구")
        self.setGeometry(100, 100, 1200, 800)

        # 중앙 위젯
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(5, 5, 5, 5)  # 마진 최소화
        layout.setSpacing(3)  # 요소 간 간격 축소

        # ===== 검색 옵션 영역 =====
        options_layout = QHBoxLayout()

        # 경로 선택
        options_layout.addWidget(QLabel("경로:"))
        self.path_input = QLineEdit()
        self.path_input.setText(str(Path.cwd()))
        self.path_input.setMaximumHeight(25)
        options_layout.addWidget(self.path_input)

        browse_btn = QPushButton("찾아보기")
        browse_btn.clicked.connect(self.browse_path)
        browse_btn.setMaximumHeight(25)
        options_layout.addWidget(browse_btn)

        # 검색어
        options_layout.addWidget(QLabel("검색:"))
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("검색할 키워드 입력...")
        self.keyword_input.returnPressed.connect(self.search)
        self.keyword_input.setMaximumHeight(25)
        options_layout.addWidget(self.keyword_input, 2)

        search_btn = QPushButton("🔍 검색")
        search_btn.clicked.connect(self.search)
        search_btn.setMaximumHeight(25)
        options_layout.addWidget(search_btn)

        layout.addLayout(options_layout)

        # ===== 경로 제외 영역 =====
        exclude_container = QWidget()
        self.exclude_layout = QHBoxLayout(exclude_container)
        self.exclude_layout.setContentsMargins(0, 3, 0, 3)
        self.exclude_layout.setSpacing(8)

        exclude_label = QLabel("제외 폴더:")
        self.exclude_layout.addWidget(exclude_label)

        add_exclude_btn = QPushButton("+ 폴더 추가")
        add_exclude_btn.clicked.connect(self.add_exclude_folder)
        add_exclude_btn.setMaximumHeight(25)
        add_exclude_btn.setMaximumWidth(100)
        self.exclude_layout.addWidget(add_exclude_btn)

        # 제외 폴더 체크박스 저장소
        self.exclude_checkboxes = {}

        self.exclude_layout.addStretch()

        exclude_container.setMaximumHeight(30)
        layout.addWidget(exclude_container)

        # ===== 추가 옵션 영역 =====
        options2_container = QWidget()
        options2_layout = QHBoxLayout(options2_container)
        options2_layout.setContentsMargins(0, 0, 0, 0)

        self.name_only_cb = QCheckBox("파일명만")
        self.name_only_cb.setMaximumHeight(25)
        self.content_only_cb = QCheckBox("내용만")
        self.content_only_cb.setMaximumHeight(25)
        self.regex_cb = QCheckBox("정규식")
        self.regex_cb.setMaximumHeight(25)

        options2_layout.addWidget(self.name_only_cb)
        options2_layout.addWidget(self.content_only_cb)
        options2_layout.addWidget(self.regex_cb)

        options2_layout.addWidget(QLabel("스레드:"))
        self.workers_spin = QSpinBox()
        self.workers_spin.setValue(8)
        self.workers_spin.setMinimum(1)
        self.workers_spin.setMaximum(16)
        self.workers_spin.setMaximumHeight(25)
        options2_layout.addWidget(self.workers_spin)

        self.refresh_btn = QPushButton("🔄 새로고침")
        self.refresh_btn.clicked.connect(self.refresh_cache)
        self.refresh_btn.setMaximumHeight(25)
        options2_layout.addWidget(self.refresh_btn)

        self.excluded_btn = QPushButton("📁 제외된 파일")
        self.excluded_btn.clicked.connect(self.show_excluded_files)
        self.excluded_btn.setMaximumHeight(25)
        options2_layout.addWidget(self.excluded_btn)

        options2_layout.addStretch()
        options2_container.setMaximumHeight(30)

        layout.addWidget(options2_container)

        # ===== 진행바 =====
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(20)
        layout.addWidget(self.progress_bar)

        # ===== 상태 메시지 =====
        self.status_label = QLabel("준비 완료")
        self.status_label.setMaximumHeight(20)
        layout.addWidget(self.status_label)

        # ===== 탭: 결과 표시 =====
        self.tabs = QTabWidget()

        # 테이블 탭
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "파일명",
            "경로",
            "크기",
            "수정 날짜",
            "검색 단어수"
        ])
        # 동적 크기 조정 설정
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(1, 1)  # 경로: 확장 가능
        self.table.verticalHeader().setVisible(False)  # 행 번호 숨김
        self.table.setSelectionBehavior(0)  # 행 선택 모드
        self.tabs.addTab(self.table, "🗂️ 결과 (테이블)")

        # 텍스트 탭
        self.text_output = QTextEdit()
        self.text_output.setReadOnly(True)
        font = QFont("Consolas")
        font.setPointSize(9)
        self.text_output.setFont(font)
        self.tabs.addTab(self.text_output, "📋 결과 (텍스트)")

        layout.addWidget(self.tabs, 1)

        # ===== 푸터 =====
        footer_container = QWidget()
        footer_layout = QHBoxLayout(footer_container)
        footer_layout.setContentsMargins(0, 0, 0, 0)
        self.result_count = QLabel("결과: 0개")
        footer_layout.addWidget(self.result_count)
        footer_layout.addStretch()
        footer_layout.addWidget(QLabel("fsearch v2.0 - GUI Edition"))
        footer_container.setMaximumHeight(20)
        layout.addWidget(footer_container)

    def add_exclude_folder(self):
        """제외할 폴더 추가"""
        folder = QFileDialog.getExistingDirectory(self, "제외할 폴더 선택")
        if folder:
            folder_name = Path(folder).name  # 폴더 이름만 추출
            full_path = folder  # 전체 경로 저장

            # 이미 추가되었는지 확인
            if full_path not in self.exclude_checkboxes:
                # 새로운 체크박스 생성
                cb = QCheckBox(folder_name)
                cb.setChecked(True)  # 새로 추가된 폴더는 제외됨
                cb.setMaximumHeight(25)
                cb.setToolTip(full_path)  # 전체 경로를 툴팁으로 표시

                # stretch 위젯 전에 체크박스 추가
                self.exclude_layout.insertWidget(
                    self.exclude_layout.count() - 1, cb
                )
                self.exclude_checkboxes[full_path] = cb

                self.status_label.setText(f"✅ 제외 폴더 추가됨: {folder_name}")
            else:
                QMessageBox.information(self, "알림", "이미 추가된 폴더입니다.")

    def browse_path(self):
        """폴더 선택 대화창"""
        folder = QFileDialog.getExistingDirectory(self, "폴더 선택")
        if folder:
            self.path_input.setText(folder)

    def search(self):
        """검색 실행"""
        keyword = self.keyword_input.text().strip()
        if not keyword:
            QMessageBox.warning(self, "입력 오류", "검색어를 입력하세요.")
            return

        path = self.path_input.text().strip()
        if not Path(path).exists():
            QMessageBox.warning(self, "경로 오류", "경로가 존재하지 않습니다.")
            return

        # 현재 검색 단어 저장
        self.current_keyword = keyword
        self.current_regex = self.regex_cb.isChecked()

        # 검색 시작
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.table.setRowCount(0)
        self.text_output.clear()
        self.results = []

        # 체크된 제외 폴더만 수집
        ignore_dirs = {folder for folder, cb in self.exclude_checkboxes.items() if cb.isChecked()}

        self.search_worker = SearchWorker(
            keyword=keyword,
            path=path,
            ignore_dirs=list(ignore_dirs),
            name_only=self.name_only_cb.isChecked(),
            content_only=self.content_only_cb.isChecked(),
            use_regex=self.regex_cb.isChecked(),
            max_workers=self.workers_spin.value()
        )

        self.search_worker.progress.connect(self.update_progress)
        self.search_worker.result_found.connect(self.add_result_row)  # 실시간 결과 추가
        self.search_worker.finished.connect(self.search_finished)
        self.search_worker.status.connect(self.update_status)
        self.search_worker.error.connect(self.search_error)
        self.search_worker.excluded_files_updated.connect(self.update_excluded_files)

        self.search_worker.start()

    def update_progress(self, value):
        """진행바 업데이트"""
        self.progress_bar.setValue(value)

    def add_result_row(self, result):
        """테이블에 결과 행 추가 (실시간)"""
        current_row_count = self.table.rowCount()
        self.table.insertRow(current_row_count)

        # 파일 크기 포맷
        size = result['size']
        if size < 1024:
            size_str = f"{size} B"
        elif size < 1024 * 1024:
            size_str = f"{size / 1024:.1f} KB"
        else:
            size_str = f"{size / (1024 * 1024):.1f} MB"

        # 각 컬럼에 데이터 입력
        # 0: 파일명 (왼쪽 정렬)
        filename_item = QTableWidgetItem(result['filename'])
        filename_item.setToolTip(result['full_path'])
        filename_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        # 1: 경로 (왼쪽 정렬)
        path_item = QTableWidgetItem(result['folder_path'])
        path_item.setToolTip(result['folder_path'])
        path_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        # 2: 크기 (오른쪽 정렬 - 숫자)
        size_item = QTableWidgetItem(size_str)
        size_item.setToolTip(str(result['size']) + " bytes")
        size_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

        # 3: 수정 날짜 (왼쪽 정렬)
        modified_item = QTableWidgetItem(result['modified'])
        modified_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        # 4: 검색 단어수 (오른쪽 정렬 - 숫자)
        match_count = result.get('match_count', 0)
        match_count_item = QTableWidgetItem(str(match_count))
        match_count_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.table.setItem(current_row_count, 0, filename_item)
        self.table.setItem(current_row_count, 1, path_item)
        self.table.setItem(current_row_count, 2, size_item)
        self.table.setItem(current_row_count, 3, modified_item)
        self.table.setItem(current_row_count, 4, match_count_item)

        # 컬럼 너비를 내용에 맞게 자동 조정
        self.table.resizeColumnsToContents()

    def update_status(self, status):
        """상태 메시지 업데이트"""
        self.status_label.setText(status)

    def search_finished(self, results):
        """검색 완료"""
        self.results = results
        self.progress_bar.setVisible(False)

        # 테이블 셀 더블클릭 시 파일 실행
        self.table.itemDoubleClicked.connect(self.open_file)

        # 검색 단어를 포함하는 셀을 굵게 표시
        if hasattr(self, 'current_keyword'):
            keyword = self.current_keyword
            keyword_lower = keyword.lower() if not self.current_regex else keyword

            for row in range(self.table.rowCount()):
                # 파일명 (컬럼 0)과 경로 (컬럼 1) 체크
                for col in [0, 1]:
                    item = self.table.item(row, col)
                    if item:
                        text = item.text()
                        # 검색 단어 찾기
                        found = False
                        if self.current_regex:
                            try:
                                if re.search(keyword, text, re.IGNORECASE):
                                    found = True
                            except:
                                pass
                        else:
                            if keyword_lower in text.lower():
                                found = True

                        # 찾으면 bold 글꼴 적용
                        if found:
                            bold_font = QFont(item.font())
                            bold_font.setBold(True)
                            item.setFont(bold_font)

        # 검색 단어수로 정렬 (내림차순 - 큰 수부터)
        self.sort_table_by_match_count()

        # 컬럼 너비 자동 조정
        self.table.resizeColumnsToContents()

        # 텍스트 탭 업데이트
        text_output = "검색 결과:\n" + "="*100 + "\n\n"
        shown_files = set()

        for result in results:
            # 각 파일당 한 번씩만 정보 표시
            file_path = result['full_path']
            if file_path not in shown_files:
                shown_files.add(file_path)

                size = result['size']
                if size < 1024:
                    size_str = f"{size} B"
                elif size < 1024 * 1024:
                    size_str = f"{size / 1024:.1f} KB"
                else:
                    size_str = f"{size / (1024 * 1024):.1f} MB"

                text_output += f"[{result['filename']}]\n"
                text_output += f"  경로: {file_path}\n"
                text_output += f"  크기: {size_str}\n"
                text_output += f"  수정일: {result['modified']}\n"
                text_output += f"  검색 단어수: {result['match_count']}\n\n"

        self.text_output.setText(text_output)

        # 파일별 검색 결과 개수 계산
        file_counts = defaultdict(int)
        for result in results:
            file_counts[result['full_path']] += 1

        self.result_count.setText(f"결과: {len(results)}개 (파일: {len(file_counts)}개)")

    def search_error(self, error):
        """검색 오류"""
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "오류", error)
        self.status_label.setText("오류 발생")

    def refresh_cache(self):
        """캐시 새로고침"""
        QMessageBox.information(self, "캐시", "검색 시 파일 목록이 새로 수집됩니다.")

    def open_file(self, item):
        """파일 또는 폴더 열기"""
        row = self.table.row(item)
        # 해당 행의 결과에서 full_path 찾기
        if row < 0 or row >= len(self.results):
            return

        result = self.results[row]
        file_path = result.get('full_path') or result.get('path')

        if not file_path or not Path(file_path).exists():
            QMessageBox.warning(self, "오류", "파일을 찾을 수 없습니다.")
            return

        try:
            # Windows에서 파일 열기
            if sys.platform == 'win32':
                os.startfile(file_path)
            # macOS
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', file_path])
            # Linux
            else:
                subprocess.Popen(['xdg-open', file_path])

            self.status_label.setText(f"✅ 파일 열음: {Path(file_path).name}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"파일을 열 수 없습니다:\n{str(e)}")

    def update_excluded_files(self, excluded_files):
        """제외된 파일 목록 업데이트"""
        self.excluded_files = excluded_files
        # 버튼 텍스트에 제외된 파일 개수 표시
        count = len(excluded_files)
        self.excluded_btn.setText(f"📁 제외된 파일 ({count})")

    def show_excluded_files(self):
        """제외된 파일 목록 보기"""
        if not self.excluded_files:
            QMessageBox.information(self, "제외된 파일", "제외된 파일이 없습니다.\n\n먼저 검색을 실행해주세요.")
            return

        # 제외된 파일 목록을 보여주는 윈도우
        from PyQt5.QtWidgets import QDialog
        dialog = QDialog(self)
        dialog.setWindowTitle("📁 제외된 파일 목록")
        dialog.setGeometry(200, 200, 800, 600)

        layout = QVBoxLayout(dialog)

        # 통계 정보
        info_label = QLabel(f"총 {len(self.excluded_files)}개의 파일이 제외되었습니다.")
        layout.addWidget(info_label)

        # 제외된 파일 목록
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setFont(QFont("Consolas", 9))

        excluded_text = "\n".join(self.excluded_files)
        text_edit.setText(excluded_text)
        layout.addWidget(text_edit)

        # 닫기 버튼
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)

        dialog.exec_()

    def sort_table_by_match_count(self):
        """테이블을 검색 단어수로 내림차순 정렬"""
        # match_count는 4번 컬럼 (0: 파일명, 1: 경로, 2: 크기, 3: 수정날짜, 4: 검색 단어수)
        items = []
        for row in range(self.table.rowCount()):
            match_count_item = self.table.item(row, 4)
            if match_count_item:
                try:
                    match_count = int(match_count_item.text())
                    items.append((row, match_count))
                except:
                    pass

        # 검색 단어수로 내림차순 정렬
        items.sort(key=lambda x: x[1], reverse=True)

        # 테이블 행 순서 변경
        # QTableWidget은 직접적인 행 이동이 없으므로, 모든 데이터를 수집했다가 다시 표시
        table_data = []
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            table_data.append(row_data)

        # 검색 단어수로 정렬된 순서대로 테이블 다시 작성
        self.table.setRowCount(0)  # 모든 행 제거
        for row, match_count in items:
            self.table.insertRow(self.table.rowCount())
            new_row = self.table.rowCount() - 1
            for col in range(len(table_data[row])):
                item = QTableWidgetItem(table_data[row][col])
                if col == 0:  # 파일명 (왼쪽 정렬)
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                elif col == 1:  # 경로 (왼쪽 정렬)
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                elif col == 2 or col == 4:  # 크기, 검색 단어수 (오른쪽 정렬)
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:  # 수정 날짜 (왼쪽 정렬)
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(new_row, col, item)


def main():
    app = QApplication(sys.argv)
    window = FSearchGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
