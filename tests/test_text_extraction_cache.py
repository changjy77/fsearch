# -*- coding: utf-8 -*-
"""TextExtractionCache: 캐시 저장/조회, FTS 색인 판단 로직, 인덱스 사용 가능 여부."""
from fsearch_gui import _can_use_index
from conftest import make_isolated_cache


def test_save_and_load_for_prefix(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    cache.save_entries({
        r"D:\proj\a.txt": (100.0, 10, "hello world"),
        r"D:\proj\b.txt": (200.0, 20, "다른 문서"),
        r"E:\other\c.txt": (300.0, 30, "other drive"),
    })

    loaded = cache.load_for_prefix(r"D:\proj")
    assert set(loaded.keys()) == {r"D:\proj\a.txt", r"D:\proj\b.txt"}
    assert loaded[r"D:\proj\a.txt"] == (100.0, 10, "hello world")


def test_delete_entries_removes_rows(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    cache.save_entries({r"D:\x.txt": (1.0, 1, "text")})
    assert cache.row_count() == 1

    cache.delete_entries([r"D:\x.txt"])
    assert cache.row_count() == 0


def test_save_entries_overwrites_existing_path(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    cache.save_entries({r"D:\x.txt": (1.0, 1, "old text")})
    cache.save_entries({r"D:\x.txt": (2.0, 2, "new text")})

    loaded = cache.load_for_prefix(r"D:\\")
    assert loaded[r"D:\x.txt"] == (2.0, 2, "new text")
    assert cache.row_count() == 1


def test_load_matching_texts_only_returns_keyword_matches(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    cache.save_entries({
        r"D:\a.txt": (1.0, 1, "네이버 관련 문서"),
        r"D:\b.txt": (2.0, 1, "카카오 관련 문서"),
    })

    matched = cache.load_matching_texts(r"D:\\", "네이버")
    assert set(matched.keys()) == {r"D:\a.txt"}


def test_should_defer_index_below_threshold_never_defers(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    # 신규 항목이 1000건 미만이면 기존 캐시가 비어 있어도 미룰 필요 없음
    assert cache.should_defer_index(999) is False


def test_should_defer_index_requires_more_than_existing_rows(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    # 기존 캐시 500건
    cache.save_entries({f"D:\\f{i}.txt": (float(i), 1, "x") for i in range(500)})

    # 신규 1000건 이상이지만 기존(500)보다 크므로 미뤄야 함
    assert cache.should_defer_index(1000) is True
    # 신규가 1000건 미만이면 기존 행 수와 무관하게 미루지 않음
    assert cache.should_defer_index(999) is False


def test_index_needs_build_false_when_already_built(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    cache.save_entries({r"D:\a.txt": (1.0, 1, "text")})
    cache.build_index()
    assert cache.index_needs_build() is False


def test_index_needs_build_true_before_build(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    cache.save_entries({r"D:\a.txt": (1.0, 1, "text")})
    assert cache.index_needs_build() is True


def test_build_index_stop_check_aborts_without_marking_built(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    # set_progress_handler는 1000 VM 명령마다 호출되므로(fsearch_gui._install_stop_handler),
    # rebuild가 그보다 먼저 끝나버리지 않도록 데이터를 충분히 크게 만든다
    cache.save_entries({f"D:\\f{i}.txt": (float(i), 1, "본문 내용 " * 3000) for i in range(500)})

    cache.build_index(stop_check=lambda: True)

    # 중단되면 트랜잭션이 롤백되어 다음 검색에서 다시 시도되어야 한다
    assert cache.index_needs_build() is True


def test_end_deferred_index_sets_built_flag(tmp_path):
    cache = make_isolated_cache(tmp_path / "cache.db")
    cache.begin_deferred_index()
    cache.save_entries({r"D:\a.txt": (1.0, 1, "text")})
    cache.end_deferred_index()
    assert cache.index_needs_build() is False


def test_can_use_index_rejects_short_keywords():
    assert _can_use_index("ab") is False
    assert _can_use_index("abc") is True


def test_can_use_index_rejects_mixed_case_non_ascii():
    # 'é'처럼 대소문자가 있는 비ASCII 문자는 SQLite LIKE의 대소문자 무시가
    # ASCII 전용이라 인덱스 결과가 파이썬 매칭과 달라질 수 있어 폴백해야 한다
    assert _can_use_index("café") is False


def test_can_use_index_allows_korean():
    # 한글은 대소문자 구분이 없어 인덱스를 그대로 써도 안전하다
    assert _can_use_index("네이버") is True
