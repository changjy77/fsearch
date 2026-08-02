#!/usr/bin/env python3
"""
fsearch GUI - PyQt5 기반 파일 검색 도구
"""

import sys
import os
import re
import time
import json
import subprocess
import logging
import zipfile
from datetime import datetime
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


def setup_logging():
    """로깅 설정"""
    log_dir = Path.home() / ".fsearch" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"fsearch_{datetime.now().strftime('%Y%m%d')}.log"

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)


class SearchHistory:
    """검색 이력 관리 - 검색어(최대 10개), 경로(최대 5개), 제외 폴더 통계"""
    def __init__(self):
        self.config_dir = Path.home() / ".fsearch"
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.history_file = self.config_dir / "search_history.json"
        self.data = self._load_data()

    def _load_data(self):
        """저장된 데이터 로드"""
        if self.history_file.exists():
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 기존 리스트 형식 호환성 처리
                    if isinstance(data, list):
                        return {
                            'keywords': data,
                            'paths': [],
                            'excluded_stats': {}
                        }
                    elif isinstance(data, dict):
                        # 누락된 키 채우기
                        if 'paths' not in data:
                            data['paths'] = []
                        if 'excluded_stats' not in data:
                            data['excluded_stats'] = {}
                        return data
            except:
                pass
        return {
            'keywords': [],
            'paths': [],
            'excluded_stats': {}
        }

    def _save_data(self):
        """데이터 저장"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except:
            pass

    def add_keyword(self, keyword):
        """검색어 추가 (중복 제거, 최대 10개)"""
        if not keyword or not keyword.strip():
            return

        keyword = keyword.strip()
        keywords = self.data.get('keywords', [])

        # 중복 제거
        if keyword in keywords:
            keywords.remove(keyword)

        # 맨 앞에 추가
        keywords.insert(0, keyword)

        # 최대 10개 유지
        self.data['keywords'] = keywords[:10]

        self._save_data()

    def add_path(self, path):
        """검색 경로 추가 (중복 제거, 최대 5개)"""
        if not path or not path.strip():
            return

        path = path.strip()
        paths = self.data.get('paths', [])

        # 중복 제거
        if path in paths:
            paths.remove(path)

        # 맨 앞에 추가
        paths.insert(0, path)

        # 최대 5개 유지
        self.data['paths'] = paths[:5]

        self._save_data()

    def add_excluded_folder(self, folder_path):
        """제외 폴더 통계 기록"""
        if not folder_path:
            return

        stats = self.data.get('excluded_stats', {})
        stats[folder_path] = stats.get(folder_path, 0) + 1
        self.data['excluded_stats'] = stats

        self._save_data()

    def get_keywords(self):
        """모든 검색 키워드 반환"""
        return self.data.get('keywords', [])

    def get_paths(self):
        """모든 검색 경로 반환"""
        return self.data.get('paths', [])

    def get_top_excluded_folders(self, limit=5):
        """자주 제외하는 폴더 반환 (상위 N개)"""
        stats = self.data.get('excluded_stats', {})
        if not stats:
            return []
        sorted_items = sorted(stats.items(), key=lambda x: x[1], reverse=True)
        return [folder for folder, _ in sorted_items[:limit]]

    def get_excluded_stats(self):
        """모든 제외 폴더 통계 반환"""
        return self.data.get('excluded_stats', {})


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
    file_processing = pyqtSignal(str)  # 현재 처리 중인 파일 이름
    no_match_files_updated = pyqtSignal(list)  # 검색어 미포함 파일 목록
    skipped_files_count = pyqtSignal(int)  # 스킵된 파일 갯수

    def __init__(self, keyword, path, ignore_dirs, name_only, content_only, use_regex, max_workers, skip_large_files=False):
        super().__init__()
        self.keyword = keyword
        self.path = path
        self.ignore_dirs = set(ignore_dirs)
        self.name_only = name_only
        self.content_only = content_only
        self.use_regex = use_regex
        self.max_workers = max_workers
        self.skip_large_files = skip_large_files
        self.results = []
        self.excluded_files = []  # 제외된 파일 목록
        self.file_cache = {}  # 실시간 캐싱: 파일 경로 → 텍스트
        self.skipped_large_files = 0  # 스킵된 대용량 파일 갯수

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

            # 매칭된 파일들과 미매칭 파일들 분류
            matched_files = {result['full_path'] for result in results}
            no_match_files = [str(f) for f in files if str(f) not in matched_files]

            self.finished.emit(results)
            self.no_match_files_updated.emit(no_match_files)
            self.status.emit(f"✅ 검색 완료: {len(results)}개 결과")

        except Exception as e:
            self.error.emit(f"오류 발생: {str(e)}")

    def _collect_files(self) -> List[Path]:
        """파일 수집 - 지정된 파일 형식만"""
        # 허용된 파일 확장자
        allowed_extensions = {
            '.doc', '.docx',      # 워드파일
            '.hwp', '.hwpx',      # 한글파일
            '.pdf',               # PDF파일
            '.xls', '.xlsx',      # 엑셀파일
            '.txt',               # 텍스트파일
            '.html', '.htm',      # HTML파일
            '.md'                 # 마크다운파일
        }

        files = []
        self.excluded_files = []  # 초기화

        # ignore_dirs를 정규화된 Path 객체로 변환
        excluded_paths = {Path(excluded).resolve() for excluded in self.ignore_dirs}

        for root, dirs, filenames in os.walk(self.path):
            # 제외할 폴더 필터링 - 전체 경로로 비교
            dirs_to_remove = []
            for d in dirs:
                dir_path = Path(root) / d
                dir_path_resolved = dir_path.resolve()

                # 제외 폴더 또는 그 하위 폴더인지 확인
                for excluded_path in excluded_paths:
                    try:
                        # dir_path가 excluded_path의 하위인지 확인
                        dir_path_resolved.relative_to(excluded_path)
                        dirs_to_remove.append(d)
                        break
                    except ValueError:
                        # 하위가 아니면 continue
                        pass

            # 제외할 폴더 제거
            for d in dirs_to_remove:
                if d in dirs:
                    dirs.remove(d)

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
        """파일 형식별로 텍스트 추출 (캐싱 지원)"""
        file_path_str = str(file_path)

        # 캐시 확인
        if file_path_str in self.file_cache:
            return self.file_cache[file_path_str]

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

            elif ext in ['.hwp', '.hwpx']:
                # 한글 파일 (.hwp, .hwpx)
                try:
                    text = ""
                    with zipfile.ZipFile(file_path, 'r') as hwp:
                        # HWP/HWPX 파일의 XML 콘텐츠 추출
                        file_list = hwp.namelist()

                        # Contents 폴더의 Section 파일들에서 텍스트 추출
                        for name in file_list:
                            if name.startswith('Contents/Section') and name.endswith('.xml'):
                                try:
                                    xml_content = hwp.read(name).decode('utf-8', errors='ignore')
                                    # 간단한 XML 파싱 - 텍스트 태그에서 내용 추출
                                    import re
                                    # <t> 태그 사이의 텍스트 추출
                                    text_matches = re.findall(r'<t>([^<]+)</t>', xml_content)
                                    for match in text_matches:
                                        text += match + " "
                                except:
                                    pass

                    return text if text else ""
                except:
                    return ""

            else:
                return ""

        except Exception as e:
            return ""

    def _search_file(self, file_path: Path, regex):
        """단일 파일 검색"""
        results = []

        # 현재 처리 중인 파일 신호 전송
        self.file_processing.emit(f"🔍 {file_path.name}")

        # 파일 정보 수집
        try:
            stat = file_path.stat()
            file_size = stat.st_size
            mod_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat.st_mtime))
        except:
            file_size = 0
            mod_time = "Unknown"

        # 대용량 파일 스킵 (5MB 이상)
        if self.skip_large_files and file_size > 5 * 1024 * 1024:
            self.skipped_large_files += 1
            self.skipped_files_count.emit(self.skipped_large_files)
            return results

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
            '.hwp': '📗',
            '.hwpx': '📗',
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
        self.read_files = []  # 읽은 파일 목록 (누적)
        self.no_match_files = []  # 검색어 미포함 파일 목록 (누적)
        self.logger = setup_logging()  # 로깅 설정
        self.search_history = SearchHistory()  # 검색 이력 관리
        self.search_start_time = None  # 검색 시작 시간
        self.search_elapsed_time = 0  # 검색 소요 시간 (초)
        self.skipped_files_count_total = 0  # 스킵된 파일 갯수
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

        # 경로 선택 (경로 이력 드롭다운)
        options_layout.addWidget(QLabel("경로:"))
        self.path_input = QComboBox()
        self.path_input.setEditable(True)
        self.path_input.setMaximumHeight(25)

        # 저장된 경로 로드
        default_path = str(Path.cwd())
        saved_paths = self.search_history.get_paths()
        if saved_paths:
            self.path_input.addItems(saved_paths)
        self.path_input.insertItem(0, default_path)
        self.path_input.setCurrentIndex(0)

        options_layout.addWidget(self.path_input)

        browse_btn = QPushButton("찾아보기")
        browse_btn.clicked.connect(self.browse_path)
        browse_btn.setMaximumHeight(25)
        options_layout.addWidget(browse_btn)

        # 검색어 (검색 이력 드롭다운)
        options_layout.addWidget(QLabel("검색:"))
        self.keyword_input = QComboBox()
        self.keyword_input.setEditable(True)
        self.keyword_input.lineEdit().setPlaceholderText("검색할 키워드 입력...")
        self.keyword_input.lineEdit().returnPressed.connect(self.search)
        self.keyword_input.setMaximumHeight(25)

        # 검색 이력 로드
        keywords = self.search_history.get_keywords()
        if keywords:
            self.keyword_input.addItems(keywords)

        options_layout.addWidget(self.keyword_input, 2)

        search_btn = QPushButton("🔍 검색")
        search_btn.clicked.connect(self.search)
        search_btn.setMaximumHeight(25)
        options_layout.addWidget(search_btn)

        layout.addLayout(options_layout)

        # ===== 경로 제외 영역 + 자주 제외하는 폴더 통계 (같은 행) =====
        combined_container = QWidget()
        combined_layout = QHBoxLayout(combined_container)
        combined_layout.setContentsMargins(0, 1, 0, 1)
        combined_layout.setSpacing(0)

        # --- 경로 제외 영역 ---
        exclude_container = QWidget()
        self.exclude_layout = QHBoxLayout(exclude_container)
        self.exclude_layout.setContentsMargins(0, 2, 0, 2)
        self.exclude_layout.setSpacing(5)

        exclude_label = QLabel("제외 폴더:")
        self.exclude_layout.addWidget(exclude_label)

        add_exclude_btn = QPushButton("+ 폴더 추가")
        add_exclude_btn.clicked.connect(self.add_exclude_folder)
        add_exclude_btn.setMaximumHeight(25)
        add_exclude_btn.setMaximumWidth(100)
        self.exclude_layout.addWidget(add_exclude_btn)

        # 제외 폴더 체크박스 저장소
        self.exclude_checkboxes = {}

        exclude_container.setMaximumHeight(30)

        combined_layout.addWidget(exclude_container)

        # --- 자주 제외하는 폴더 통계 영역 ---
        stats_container = QWidget()
        stats_layout = QHBoxLayout(stats_container)
        stats_layout.setContentsMargins(0, 2, 0, 2)
        stats_layout.setSpacing(5)

        stats_label = QLabel("📊 자주 제외하는 폴더:")
        stats_layout.addWidget(stats_label)

        # 자주 제외하는 폴더 표시
        self.stats_label = QLabel()
        self.update_excluded_stats_display()
        stats_layout.addWidget(self.stats_label)

        stats_container.setMaximumHeight(30)

        combined_layout.addWidget(stats_container)

        # 남은 공간 채우기
        combined_layout.addStretch()

        combined_container.setMaximumHeight(30)
        layout.addWidget(combined_container)

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
        self.skip_large_cb = QCheckBox("대용량파일 스킵(>5MB)")
        self.skip_large_cb.setMaximumHeight(25)

        options2_layout.addWidget(self.name_only_cb)
        options2_layout.addWidget(self.content_only_cb)
        options2_layout.addWidget(self.regex_cb)
        options2_layout.addWidget(self.skip_large_cb)

        options2_layout.addWidget(QLabel("스레드:"))
        self.workers_spin = QSpinBox()
        self.workers_spin.setValue(8)
        self.workers_spin.setMinimum(1)
        self.workers_spin.setMaximum(32)
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

        self.read_files_btn = QPushButton("🔍 찾은 파일 수 (0)")
        self.read_files_btn.clicked.connect(self.show_read_files)
        self.read_files_btn.setMaximumHeight(25)
        options2_layout.addWidget(self.read_files_btn)

        self.no_match_files_btn = QPushButton("❌ 검색어 미포함 파일 수 (0)")
        self.no_match_files_btn.clicked.connect(self.show_no_match_files)
        self.no_match_files_btn.setMaximumHeight(25)
        options2_layout.addWidget(self.no_match_files_btn)

        self.performance_btn = QPushButton("⏱️ 완료시간 (0.00초)")
        self.performance_btn.setMaximumHeight(25)
        self.performance_btn.setEnabled(False)
        options2_layout.addWidget(self.performance_btn)

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

        # 탭 스타일 설정
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ddd;
            }
            QTabBar::tab {
                background-color: #f5f5f5;
                color: #333;
                padding: 6px 16px;
                margin-right: 2px;
                border: 1px solid #ddd;
                border-bottom: none;
                font-size: 10pt;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                color: #0078d4;
                border: 1px solid #ddd;
                border-bottom: 3px solid #0078d4;
                font-weight: bold;
            }
            QTabBar::tab:hover:!selected {
                background-color: #efefef;
            }
        """)

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

    def update_excluded_stats_display(self):
        """자주 제외하는 폴더 통계 표시"""
        top_folders = self.search_history.get_top_excluded_folders(limit=3)
        stats = self.search_history.get_excluded_stats()

        if not top_folders:
            self.stats_label.setText("통계 없음")
            return

        # 상위 3개 폴더와 사용 횟수 표시
        stats_text = " | ".join(
            [f"{Path(folder).name} ({stats.get(folder, 0)}회)" for folder in top_folders]
        )
        self.stats_label.setText(stats_text)

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

                # 제외 폴더 통계 기록
                self.search_history.add_excluded_folder(full_path)

                # 통계 표시 업데이트
                self.update_excluded_stats_display()

                self.status_label.setText(f"✅ 제외 폴더 추가됨: {folder_name}")
            else:
                QMessageBox.information(self, "알림", "이미 추가된 폴더입니다.")

    def browse_path(self):
        """폴더 선택 대화창"""
        folder = QFileDialog.getExistingDirectory(self, "폴더 선택")
        if folder:
            self.path_input.setCurrentText(folder)
            self.search_history.add_path(folder)

    def search(self):
        """검색 실행"""
        keyword = self.keyword_input.currentText().strip()
        if not keyword:
            QMessageBox.warning(self, "입력 오류", "검색어를 입력하세요.")
            return

        path = self.path_input.currentText().strip()
        if not Path(path).exists():
            QMessageBox.warning(self, "경로 오류", "경로가 존재하지 않습니다.")
            return

        # 검색 이력에 저장 (최대 10개)
        self.search_history.add_keyword(keyword)

        # 검색 경로 저장 (최대 5개)
        self.search_history.add_path(path)

        # 현재 검색 단어 저장
        self.current_keyword = keyword
        self.current_regex = self.regex_cb.isChecked()

        # 검색 시작 시간 기록
        self.search_start_time = time.time()
        self.performance_btn.setText("⏱️ 완료시간 (진행중...)")
        self.performance_btn.setEnabled(False)

        # 스킵된 파일 갯수 초기화
        self.skipped_files_count_total = 0

        # 검색 시작 - 이전 결과 모두 초기화
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.table.setRowCount(0)
        self.text_output.clear()
        self.results = []

        # 이전 검색 결과 초기화
        self.excluded_files = []
        self.read_files = []
        self.no_match_files = []

        # 관련 버튼 초기화
        self.excluded_btn.setText("❌ 제외된 파일 (0)")
        self.read_files_btn.setText("🔍 찾은 파일 수 (0)")
        self.no_match_files_btn.setText("❌ 검색어 미포함 파일 수 (0)")

        # 체크된 제외 폴더만 수집
        ignore_dirs = {folder for folder, cb in self.exclude_checkboxes.items() if cb.isChecked()}

        self.search_worker = SearchWorker(
            keyword=keyword,
            path=path,
            ignore_dirs=list(ignore_dirs),
            name_only=self.name_only_cb.isChecked(),
            content_only=self.content_only_cb.isChecked(),
            use_regex=self.regex_cb.isChecked(),
            max_workers=self.workers_spin.value(),
            skip_large_files=self.skip_large_cb.isChecked()
        )

        self.search_worker.progress.connect(self.update_progress)
        self.search_worker.result_found.connect(self.add_result_row)  # 실시간 결과 추가
        self.search_worker.finished.connect(self.search_finished)
        self.search_worker.status.connect(self.update_status)
        self.search_worker.error.connect(self.search_error)
        self.search_worker.excluded_files_updated.connect(self.update_excluded_files)
        self.search_worker.file_processing.connect(self.update_current_file)  # 현재 처리 중인 파일
        self.search_worker.no_match_files_updated.connect(self.update_no_match_files)  # 검색어 미포함 파일
        self.search_worker.skipped_files_count.connect(self.update_skipped_files_count)  # 스킵된 파일 갯수

        # 로깅
        self.logger.info(f"검색 시작 - 경로: {path}, 검색어: {keyword}")

        self.search_worker.start()

    def update_progress(self, value):
        """진행바 업데이트"""
        self.progress_bar.setValue(value)

    def update_skipped_files_count(self, count):
        """스킵된 파일 갯수 업데이트"""
        self.skipped_files_count_total = count

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
        # 0: 파일명 (검색어 강조 표시)
        filename_text = result['filename']
        if hasattr(self, 'current_keyword') and self.current_keyword:
            highlighted_filename = self._highlight_keyword(filename_text, self.current_keyword, self.current_regex)
            if '<span' in highlighted_filename:
                filename_label = QLabel()
                filename_label.setText(highlighted_filename)
                filename_label.setStyleSheet("padding: 3px;")
                filename_label.setToolTip(result['full_path'])
                self.table.setCellWidget(current_row_count, 0, filename_label)
            else:
                filename_item = QTableWidgetItem(filename_text)
                filename_item.setToolTip(result['full_path'])
                filename_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(current_row_count, 0, filename_item)
        else:
            filename_item = QTableWidgetItem(filename_text)
            filename_item.setToolTip(result['full_path'])
            filename_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.table.setItem(current_row_count, 0, filename_item)

        # 1: 경로 (검색어 강조 표시)
        path_text = result['folder_path']
        if hasattr(self, 'current_keyword') and self.current_keyword:
            highlighted_path = self._highlight_keyword(path_text, self.current_keyword, self.current_regex)
            if '<span' in highlighted_path:
                path_label = QLabel()
                path_label.setText(highlighted_path)
                path_label.setStyleSheet("padding: 3px;")
                path_label.setToolTip(path_text)
                self.table.setCellWidget(current_row_count, 1, path_label)
            else:
                path_item = QTableWidgetItem(path_text)
                path_item.setToolTip(path_text)
                path_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(current_row_count, 1, path_item)
        else:
            path_item = QTableWidgetItem(path_text)
            path_item.setToolTip(path_text)
            path_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.table.setItem(current_row_count, 1, path_item)

        # 2: 크기 (오른쪽 정렬 - 숫자)
        size_item = QTableWidgetItem(size_str)
        size_item.setToolTip(str(result['size']) + " bytes")
        size_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.table.setItem(current_row_count, 2, size_item)

        # 3: 수정 날짜 (왼쪽 정렬)
        modified_item = QTableWidgetItem(result['modified'])
        modified_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.table.setItem(current_row_count, 3, modified_item)

        # 4: 검색 단어수 (오른쪽 정렬 - 숫자)
        match_count = result.get('match_count', 0)
        match_count_item = QTableWidgetItem(str(match_count))
        match_count_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.table.setItem(current_row_count, 4, match_count_item)

        # 컬럼 너비를 내용에 맞게 자동 조정
        self.table.resizeColumnsToContents()

    def update_status(self, status):
        """상태 메시지 업데이트"""
        self.status_label.setText(status)

    def update_current_file(self, file_info):
        """현재 처리 중인 파일 표시"""
        self.status_label.setText(f"검색 중: {file_info}")

    def update_no_match_files(self, no_match_files):
        """검색어 미포함 파일 누적 업데이트"""
        # 중복 제거하면서 누적
        for file_path in no_match_files:
            if file_path not in self.no_match_files:
                self.no_match_files.append(file_path)

        # 버튼 텍스트 업데이트
        self.no_match_files_btn.setText(f"❌ 검색어 미포함 파일 수 ({len(self.no_match_files)})")

    def search_finished(self, results):
        """검색 완료"""
        # 완료 시간 계산
        if self.search_start_time:
            self.search_elapsed_time = time.time() - self.search_start_time
            self.performance_btn.setText(f"⏱️ 완료시간 ({self.search_elapsed_time:.2f}초)")
            self.performance_btn.setEnabled(True)

        self.results = results
        self.progress_bar.setVisible(False)

        # 테이블 셀 더블클릭 시 파일 실행
        self.table.itemDoubleClicked.connect(self.open_file)

        # 검색 단어를 포함하는 셀을 굵게 표시 + 빨간색
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

                        # 찾으면 bold 글꼴 + 빨간색 적용
                        if found:
                            bold_font = QFont(item.font())
                            bold_font.setBold(True)
                            item.setFont(bold_font)
                            item.setForeground(QColor('red'))

        # 검색 단어수로 정렬 (내림차순 - 큰 수부터)
        self.sort_table_by_match_count()

        # 컬럼 너비 자동 조정
        self.table.resizeColumnsToContents()

        # 텍스트 탭 업데이트 (검색어 굵게 표기)
        text_output = "<html><body><pre>검색 결과:\n" + "="*100 + "\n\n"
        shown_files = set()

        # 검색어 준비
        keyword = getattr(self, 'current_keyword', '')
        use_regex = getattr(self, 'current_regex', False)

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

                # 파일명에서 검색어를 굵게 표기
                filename = result['filename']
                highlighted_filename = self._highlight_keyword(filename, keyword, use_regex)

                # 경로에서 검색어를 굵게 표기
                highlighted_path = self._highlight_keyword(file_path, keyword, use_regex)

                text_output += f"[{highlighted_filename}]\n"
                text_output += f"  경로: {highlighted_path}\n"
                text_output += f"  크기: {size_str}\n"
                text_output += f"  수정일: {result['modified']}\n"
                text_output += f"  검색 단어수: {result['match_count']}\n\n"

        text_output += "</pre></body></html>"
        self.text_output.setHtml(text_output)

        # 파일별 검색 결과 개수 계산
        file_counts = defaultdict(int)
        for result in results:
            file_counts[result['full_path']] += 1

        self.result_count.setText(f"결과: {len(results)}개 (파일: {len(file_counts)}개)")

        # 읽은 파일 누적 (중복 제거)
        for file_path in file_counts.keys():
            if file_path not in self.read_files:
                self.read_files.append(file_path)

        # 읽은 파일 버튼 업데이트
        self.read_files_btn.setText(f"🔍 찾은 파일 수 ({len(self.read_files)})")

        # 상태 메시지 업데이트 - 스킵된 파일 정보 포함
        skip_info = f" (스킵: {self.skipped_files_count_total}개 대용량파일)" if self.skipped_files_count_total > 0 else ""
        self.status_label.setText(f"✅ 검색 완료: {len(results)}개 결과{skip_info}")

        # 로깅 - 검색 결과
        self.logger.info(f"검색 완료 - 총 {len(results)}개 결과 (파일: {len(file_counts)}개, 스킵: {self.skipped_files_count_total}개)")

    def _highlight_keyword(self, text: str, keyword: str, use_regex: bool) -> str:
        """텍스트에서 검색어를 HTML 굵게 + 빨간색으로 표기"""
        if not keyword:
            return text

        if use_regex:
            try:
                # 정규식 사용
                pattern = re.compile(f'({keyword})', re.IGNORECASE)
                return pattern.sub(r'<span style="color: red; font-weight: bold;">\1</span>', text)
            except:
                return text
        else:
            # 일반 문자열 검색 (대소문자 무시)
            # 모든 일치 부분을 굵게 + 빨간색으로 표기
            pattern = re.compile(f'({re.escape(keyword)})', re.IGNORECASE)
            return pattern.sub(r'<span style="color: red; font-weight: bold;">\1</span>', text)

    def search_error(self, error):
        """검색 오류"""
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "오류", error)
        self.status_label.setText("오류 발생")
        # 로깅
        self.logger.error(f"검색 오류: {error}")

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

            filename = Path(file_path).name
            self.status_label.setText(f"✅ 파일 열음: {filename}")

            # 로깅
            self.logger.info(f"파일 실행 - {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"파일을 열 수 없습니다:\n{str(e)}")
            self.logger.error(f"파일 실행 오류 - {file_path}: {str(e)}")

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

    def show_read_files(self):
        """읽은 파일 목록 보기"""
        if not self.read_files:
            QMessageBox.information(self, "찾은 파일 수", "찾은 파일이 없습니다.\n\n먼저 검색을 실행해주세요.")
            return

        # 읽은 파일 목록을 보여주는 윈도우
        from PyQt5.QtWidgets import QDialog
        dialog = QDialog(self)
        dialog.setWindowTitle("🔍 찾은 파일 수")
        dialog.setGeometry(200, 200, 800, 600)

        layout = QVBoxLayout(dialog)

        # 통계 정보
        info_label = QLabel(f"검색어를 찾은 파일 {len(self.read_files)}개 (누적)")
        layout.addWidget(info_label)

        # 읽은 파일 목록
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setFont(QFont("Consolas", 9))

        read_text = "\n".join(self.read_files)
        text_edit.setText(read_text)
        layout.addWidget(text_edit)

        # 닫기 버튼
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)

        dialog.exec_()

    def show_no_match_files(self):
        """검색어 미포함 파일 목록 보기"""
        if not self.no_match_files:
            QMessageBox.information(self, "검색어 미포함 파일 수", "검색어 미포함 파일이 없습니다.\n\n먼저 검색을 실행해주세요.")
            return

        # 검색어 미포함 파일 목록을 보여주는 윈도우
        from PyQt5.QtWidgets import QDialog
        dialog = QDialog(self)
        dialog.setWindowTitle("❌ 검색어 미포함 파일 수")
        dialog.setGeometry(200, 200, 800, 600)

        layout = QVBoxLayout(dialog)

        # 통계 정보
        info_label = QLabel(f"검색어를 찾지 못한 파일 {len(self.no_match_files)}개 (누적)")
        layout.addWidget(info_label)

        # 검색어 미포함 파일 목록
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setFont(QFont("Consolas", 9))

        no_match_text = "\n".join(self.no_match_files)
        text_edit.setText(no_match_text)
        layout.addWidget(text_edit)

        # 닫기 버튼
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)

        dialog.exec_()

    def sort_table_by_match_count(self):
        """테이블을 검색 단어수로 내림차순 정렬"""
        # results를 match_count로 정렬
        sorted_results = sorted(self.results, key=lambda x: x.get('match_count', 0), reverse=True)

        # 테이블 초기화
        self.table.setRowCount(0)

        # 정렬된 결과를 테이블에 다시 추가 (강조 표기 포함)
        shown_files = set()
        for result in sorted_results:
            file_path = result['full_path']
            if file_path not in shown_files:
                shown_files.add(file_path)

                # add_result_row와 동일한 로직으로 추가
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
                # 0: 파일명 (검색어 강조 표시)
                filename_text = result['filename']
                if hasattr(self, 'current_keyword') and self.current_keyword:
                    highlighted_filename = self._highlight_keyword(filename_text, self.current_keyword, self.current_regex)
                    if '<span' in highlighted_filename:
                        filename_label = QLabel()
                        filename_label.setText(highlighted_filename)
                        filename_label.setStyleSheet("padding: 3px;")
                        filename_label.setToolTip(result['full_path'])
                        self.table.setCellWidget(current_row_count, 0, filename_label)
                    else:
                        filename_item = QTableWidgetItem(filename_text)
                        filename_item.setToolTip(result['full_path'])
                        filename_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                        self.table.setItem(current_row_count, 0, filename_item)
                else:
                    filename_item = QTableWidgetItem(filename_text)
                    filename_item.setToolTip(result['full_path'])
                    filename_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    self.table.setItem(current_row_count, 0, filename_item)

                # 1: 경로 (검색어 강조 표시)
                path_text = result['folder_path']
                if hasattr(self, 'current_keyword') and self.current_keyword:
                    highlighted_path = self._highlight_keyword(path_text, self.current_keyword, self.current_regex)
                    if '<span' in highlighted_path:
                        path_label = QLabel()
                        path_label.setText(highlighted_path)
                        path_label.setStyleSheet("padding: 3px;")
                        path_label.setToolTip(path_text)
                        self.table.setCellWidget(current_row_count, 1, path_label)
                    else:
                        path_item = QTableWidgetItem(path_text)
                        path_item.setToolTip(path_text)
                        path_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                        self.table.setItem(current_row_count, 1, path_item)
                else:
                    path_item = QTableWidgetItem(path_text)
                    path_item.setToolTip(path_text)
                    path_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    self.table.setItem(current_row_count, 1, path_item)

                # 2: 크기
                size_item = QTableWidgetItem(size_str)
                size_item.setToolTip(str(result['size']) + " bytes")
                size_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(current_row_count, 2, size_item)

                # 3: 수정 날짜
                modified_item = QTableWidgetItem(result['modified'])
                modified_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.table.setItem(current_row_count, 3, modified_item)

                # 4: 검색 단어수
                match_count = result.get('match_count', 0)
                match_count_item = QTableWidgetItem(str(match_count))
                match_count_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(current_row_count, 4, match_count_item)


def main():
    app = QApplication(sys.argv)
    window = FSearchGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
