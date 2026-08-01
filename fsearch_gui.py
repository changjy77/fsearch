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
    QSplitter, QMessageBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon, QColor, QFont, QCursor


class SearchWorker(QThread):
    """검색을 별도 스레드에서 실행"""
    progress = pyqtSignal(int)
    result_found = pyqtSignal(dict)  # 결과를 실시간으로 전송
    finished = pyqtSignal(list)
    status = pyqtSignal(str)
    error = pyqtSignal(str)

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
        """파일 수집"""
        files = []
        for root, dirs, filenames in os.walk(self.path):
            dirs[:] = [d for d in dirs if d not in self.ignore_dirs]
            for filename in filenames:
                file_path = Path(root) / filename
                files.append(file_path)
        return files

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
        content_matches = []  # (줄번호, 내용) 임시 저장

        # 파일명 검색
        if not self.content_only:
            if self._match_keyword(file_path.name, regex):
                total_match_count += 1

        # 파일 내용 검색 (먼저 카운트만)
        if not self.name_only and not self._is_binary(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    for line_num, line in enumerate(f, 1):
                        if self._match_keyword(line, regex):
                            total_match_count += 1
                            content_matches.append((line_num, line.rstrip()[:80]))
            except:
                pass

        # 파일명 매칭 결과 추가 (총 카운트 포함)
        if not self.content_only and total_match_count > 0:
            if self._match_keyword(file_path.name, regex):
                results.append({
                    'type': 'filename',
                    'filename': filename_with_icon,
                    'folder_path': folder_path,
                    'full_path': str(file_path),
                    'size': file_size,
                    'modified': mod_time,
                    'line': None,
                    'content': None,
                    'match_count': total_match_count  # 파일 내 총 매칭 횟수
                })

        # 파일 내용 매칭 결과 추가 (총 카운트 포함)
        for line_num, content in content_matches:
            results.append({
                'type': 'content',
                'filename': filename_with_icon,
                'folder_path': folder_path,
                'full_path': str(file_path),
                'size': file_size,
                'modified': mod_time,
                'line': line_num,
                'content': content,
                'match_count': total_match_count  # 파일 내 총 매칭 횟수
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
        self.init_ui()

    def init_ui(self):
        """UI 초기화"""
        self.setWindowTitle("🔍 fsearch - 파일 검색 도구")
        self.setGeometry(100, 100, 1200, 800)

        # 중앙 위젯
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # ===== 검색 옵션 영역 =====
        options_layout = QHBoxLayout()

        # 경로 선택
        options_layout.addWidget(QLabel("경로:"))
        self.path_input = QLineEdit()
        self.path_input.setText(str(Path.cwd()))
        options_layout.addWidget(self.path_input)

        browse_btn = QPushButton("찾아보기")
        browse_btn.clicked.connect(self.browse_path)
        options_layout.addWidget(browse_btn)

        # 검색어
        options_layout.addWidget(QLabel("검색:"))
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("검색할 키워드 입력...")
        self.keyword_input.returnPressed.connect(self.search)
        options_layout.addWidget(self.keyword_input, 2)

        search_btn = QPushButton("🔍 검색")
        search_btn.clicked.connect(self.search)
        options_layout.addWidget(search_btn)

        layout.addLayout(options_layout)

        # ===== 추가 옵션 영역 =====
        options2_layout = QHBoxLayout()

        self.name_only_cb = QCheckBox("파일명만")
        self.content_only_cb = QCheckBox("내용만")
        self.regex_cb = QCheckBox("정규식")

        options2_layout.addWidget(self.name_only_cb)
        options2_layout.addWidget(self.content_only_cb)
        options2_layout.addWidget(self.regex_cb)

        options2_layout.addWidget(QLabel("스레드:"))
        self.workers_spin = QSpinBox()
        self.workers_spin.setValue(8)
        self.workers_spin.setMinimum(1)
        self.workers_spin.setMaximum(16)
        options2_layout.addWidget(self.workers_spin)

        self.refresh_btn = QPushButton("🔄 새로고침")
        self.refresh_btn.clicked.connect(self.refresh_cache)
        options2_layout.addWidget(self.refresh_btn)

        options2_layout.addStretch()

        layout.addLayout(options2_layout)

        # ===== 진행바 =====
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # ===== 상태 메시지 =====
        self.status_label = QLabel("준비 완료")
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
        footer_layout = QHBoxLayout()
        self.result_count = QLabel("결과: 0개")
        footer_layout.addWidget(self.result_count)
        footer_layout.addStretch()
        footer_layout.addWidget(QLabel("fsearch v1.1 - GUI Edition"))
        layout.addLayout(footer_layout)

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

        # 검색 시작
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.table.setRowCount(0)
        self.text_output.clear()
        self.results = []

        ignore_dirs = {'.git', '__pycache__', 'node_modules', '.venv', 'venv'}

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


def main():
    app = QApplication(sys.argv)
    window = FSearchGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
