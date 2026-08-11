# -*- coding: utf-8 -*-
"""SearchWorker: 검색 정확성, MAX_RESULTS 상한(오버슈트/결과유실 방지), stop_flag 반응성.

프로세스풀(ProcessPoolExecutor) 스폰은 이 저장소를 다루던 세션에서 반복적으로
환경 의존적 지연을 겪었으므로, worth_pool 임계값(8개) 미만으로 대상을 유지하거나
캐시를 미리 채워 추출 자체를 건너뛰는 방식으로 프로세스풀을 우회한다.
"""
import sqlite3
import zipfile

from conftest import make_isolated_cache, make_search_worker


def run_worker_sync(worker):
    """QThread.start() 대신 run()을 현재 스레드에서 동기 실행해 결과를 즉시 받는다."""
    holder = {}
    worker.finished.connect(lambda results: holder.__setitem__('results', results))
    worker.run()
    return holder.get('results', [])


def test_content_search_finds_matching_files(qapp, tmp_path, isolated_cache):
    (tmp_path / "a.txt").write_text("네이버 관련 내용", encoding="utf-8")
    (tmp_path / "b.txt").write_text("카카오 관련 내용", encoding="utf-8")

    worker = make_search_worker("네이버", tmp_path, isolated_cache)
    results = run_worker_sync(worker)

    assert len(results) == 1
    assert results[0]['full_path'].endswith("a.txt")


def test_filename_only_search(qapp, tmp_path, isolated_cache):
    (tmp_path / "report_2024.txt").write_text("아무 내용", encoding="utf-8")
    (tmp_path / "other.txt").write_text("아무 내용", encoding="utf-8")

    worker = make_search_worker("report", tmp_path, isolated_cache, name_only=True)
    results = run_worker_sync(worker)

    assert len(results) == 1
    assert "report_2024.txt" in results[0]['full_path']


def test_no_match_returns_empty(qapp, tmp_path, isolated_cache):
    (tmp_path / "a.txt").write_text("무관한 내용", encoding="utf-8")

    worker = make_search_worker("존재하지않는키워드", tmp_path, isolated_cache)
    results = run_worker_sync(worker)

    assert results == []


def _prime_cache_with_zip(cache_db_path, zip_path, n_entries, keyword):
    """zip 내부에 매치를 다량(n_entries) 미리 캐시에 채워, 검색 시 프로세스풀
    사전 추출 없이 곧바로 검색 단계(MAX_RESULTS 로직)로 가게 한다."""
    entries_text = {}
    with zipfile.ZipFile(zip_path, 'w') as zf:
        for i in range(n_entries):
            text = f"본문 내용에 {keyword} 언급 {i}"
            name = f"entry_{i}.txt"
            zf.writestr(name, text)
            entries_text[name] = text

    conn = sqlite3.connect(str(cache_db_path))
    zip_stat = zip_path.stat()
    zip_str = str(zip_path)
    conn.execute("INSERT INTO text_cache VALUES (?,?,?,?)", (zip_str, zip_stat.st_mtime, zip_stat.st_size, ""))
    with zipfile.ZipFile(zip_path) as zf:
        for info in zf.infolist():
            key = f"{zip_str}!{info.filename}"
            text = entries_text[info.filename]
            conn.execute("INSERT INTO text_cache VALUES (?,?,?,?)", (key, info.CRC, info.file_size, text))
    conn.execute("INSERT INTO text_fts(text_fts) VALUES('rebuild')")
    conn.execute("INSERT OR REPLACE INTO cache_meta (key, value) VALUES ('fts_built', '1')")
    conn.commit()
    conn.close()


def test_max_results_stops_without_overshoot_or_loss(qapp, tmp_path):
    """실제로 발견됐던 회귀 버그의 재현 테스트:
    zip 파일 하나에 매치가 몰려 있으면, 상한(MAX_RESULTS)에 도달한 순간
    (1) zip 전체를 다 처리할 때까지 오버슈트하지 않아야 하고
    (2) 이미 진행 중이던 작업의 결과가 통째로 유실되지도 않아야 한다."""
    cache_db = tmp_path / "cache.db"
    cache = make_isolated_cache(cache_db)

    zip_path = tmp_path / "big.zip"
    _prime_cache_with_zip(cache_db, zip_path, n_entries=3000, keyword=".doc")

    worker = make_search_worker(".doc", tmp_path, cache, max_results=50)
    results = run_worker_sync(worker)

    assert len(results) == 50
    assert worker._max_results_hit is True


def test_max_results_not_triggered_when_under_limit(qapp, tmp_path):
    cache_db = tmp_path / "cache.db"
    cache = make_isolated_cache(cache_db)

    zip_path = tmp_path / "small.zip"
    _prime_cache_with_zip(cache_db, zip_path, n_entries=10, keyword=".doc")

    worker = make_search_worker(".doc", tmp_path, cache, max_results=50)
    results = run_worker_sync(worker)

    assert len(results) == 10
    assert worker._max_results_hit is False


def test_stop_flag_set_before_run_yields_empty_results(qapp, tmp_path, isolated_cache):
    """이미 중단 신호가 켜진 상태로 검색을 시작하면, 검색 루프에 진입하는 즉시
    멈추고 안전하게 빈 결과로 끝나야 한다(진행 중이던 검색 없이 크래시 없음 확인)."""
    (tmp_path / "a.txt").write_text("매치될 내용", encoding="utf-8")

    worker = make_search_worker("매치", tmp_path, isolated_cache)
    worker.stop_flag = True
    results = run_worker_sync(worker)

    assert results == []
