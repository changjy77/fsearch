# -*- coding: utf-8 -*-
"""공용 픽스처. 실제 ~/.fsearch 캐시/이력 파일은 절대 건드리지 않고, 매 테스트마다
격리된 임시 DB/디렉토리를 사용한다 (__new__로 생성해 __init__의 홈 디렉토리 접근을 우회)."""
import sqlite3
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from PyQt5.QtWidgets import QApplication

from fsearch_gui import SearchHistory, SearchWorker, TextExtractionCache


@pytest.fixture(scope="session")
def qapp():
    """PyQt 위젯/스레드/타이머는 QApplication 인스턴스가 있어야 정상 동작한다."""
    app = QApplication.instance() or QApplication([])
    return app


def make_isolated_cache(db_path: Path) -> TextExtractionCache:
    """실제 캐시 생성자(Path.home() 접근)를 우회해 격리된 DB를 가리키는
    TextExtractionCache를 만든다."""
    cache = TextExtractionCache.__new__(TextExtractionCache)
    cache.db_path = db_path
    conn = sqlite3.connect(str(db_path))
    conn.execute(
        "CREATE TABLE IF NOT EXISTS text_cache "
        "(path TEXT PRIMARY KEY, mtime REAL, size INTEGER, text TEXT)"
    )
    conn.execute("CREATE TABLE IF NOT EXISTS cache_meta (key TEXT PRIMARY KEY, value TEXT)")
    conn.execute(
        "CREATE VIRTUAL TABLE IF NOT EXISTS text_fts USING fts5("
        "text, content='text_cache', content_rowid='rowid', tokenize='trigram')"
    )
    TextExtractionCache._create_triggers(conn)
    conn.commit()
    conn.close()
    cache.fts_available = True
    cache.cleared_by_version = 0
    return cache


@pytest.fixture
def isolated_cache(tmp_path):
    """격리된 SQLite 캐시(text_cache/cache_meta/text_fts)."""
    return make_isolated_cache(tmp_path / "isolated_cache.db")


def make_isolated_history(history_file: Path) -> SearchHistory:
    """실제 검색 이력 파일(Path.home() 접근)을 우회해 격리된 JSON을 쓰는
    SearchHistory를 만든다."""
    history = SearchHistory.__new__(SearchHistory)
    history.history_file = history_file
    history.data = {'keywords': [], 'paths': [], 'excluded_stats': {}}
    return history


@pytest.fixture
def isolated_history(tmp_path):
    return make_isolated_history(tmp_path / "search_history.json")


def make_search_worker(keyword, path, cache, max_results=None, **kwargs):
    """실제 캐시 대신 격리된 캐시를 쓰는 SearchWorker를 만든다.
    SearchWorker() 생성자가 실제 ~/.fsearch/text_cache.db에 잠깐 접속하지만
    CREATE TABLE IF NOT EXISTS만 실행하므로 실제 캐시에는 영향이 없고,
    곧바로 text_cache를 격리된 것으로 교체한다."""
    defaults = dict(
        ignore_dirs=[], name_only=False, content_only=False,
        use_regex=False, max_workers=4, skip_large_files=False,
    )
    defaults.update(kwargs)
    worker = SearchWorker(keyword=keyword, path=str(path), **defaults)
    worker.text_cache = cache
    if max_results is not None:
        worker.MAX_RESULTS = max_results
    return worker
