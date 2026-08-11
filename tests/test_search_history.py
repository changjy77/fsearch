# -*- coding: utf-8 -*-
"""SearchHistory: 검색어/경로 이력, 제외 폴더 통계."""
from conftest import make_isolated_history


def test_add_path_dedup_and_max_five(tmp_path):
    history = make_isolated_history(tmp_path / "history.json")
    for p in [r"C:\a", r"C:\b", r"C:\c", r"C:\d", r"C:\e", r"C:\f"]:
        history.add_path(p)

    assert history.get_paths() == [r"C:\f", r"C:\e", r"C:\d", r"C:\c", r"C:\b"]

    history.add_path(r"C:\c")
    result = history.get_paths()
    assert result[0] == r"C:\c"
    assert result.count(r"C:\c") == 1
    assert len(result) == 5


def test_add_keyword_dedup_and_max_ten(tmp_path):
    history = make_isolated_history(tmp_path / "history.json")
    for i in range(12):
        history.add_keyword(f"kw{i}")

    keywords = history.get_keywords()
    assert len(keywords) == 10
    assert keywords[0] == "kw11"
    assert "kw0" not in keywords
    assert "kw1" not in keywords


def test_add_keyword_ignores_blank(tmp_path):
    history = make_isolated_history(tmp_path / "history.json")
    history.add_keyword("   ")
    history.add_keyword("")
    assert history.get_keywords() == []


def test_clear_paths_and_excluded_stats(tmp_path):
    history = make_isolated_history(tmp_path / "history.json")
    history.add_path(r"D:\x")
    history.add_excluded_folder(r"D:\AppData")

    history.clear_paths()
    assert history.get_paths() == []
    # 경로 초기화는 제외 폴더 통계와 별개 메서드이므로 아직 남아 있어야 한다
    assert history.get_excluded_stats() == {r"D:\AppData": 1}

    history.clear_excluded_stats()
    assert history.get_excluded_stats() == {}


def test_excluded_folder_stats_increment_and_top_n(tmp_path):
    history = make_isolated_history(tmp_path / "history.json")
    for _ in range(3):
        history.add_excluded_folder(r"D:\AppData")
    for _ in range(1):
        history.add_excluded_folder(r"D:\Temp")

    stats = history.get_excluded_stats()
    assert stats[r"D:\AppData"] == 3
    assert stats[r"D:\Temp"] == 1
    assert history.get_top_excluded_folders(limit=1) == [r"D:\AppData"]


def test_data_persists_to_disk(tmp_path):
    history_file = tmp_path / "history.json"
    history = make_isolated_history(history_file)
    history.add_path(r"E:\project")

    # 새 인스턴스로 다시 열어도 방금 저장한 값이 로드되는지 확인
    reloaded = make_isolated_history(history_file)
    reloaded.data = reloaded._load_data()
    assert reloaded.get_paths() == [r"E:\project"]
