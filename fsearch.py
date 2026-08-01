#!/usr/bin/env python3
"""
빠른 파일 검색 도구
파일명과 파일 내용을 동시에 검색합니다.
"""

import argparse
import os
import re
import sys
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Tuple, Optional


class FileSearcher:
    def __init__(self, keyword: str, path: Path, ignore_dirs: List[str], use_regex: bool, max_workers: int = 8):
        self.keyword = keyword
        self.path = path
        self.ignore_dirs = set(ignore_dirs)
        self.use_regex = use_regex
        self.max_workers = max_workers
        self.results = []

        # 정규식 컴파일
        if use_regex:
            try:
                self.regex = re.compile(keyword, re.IGNORECASE)
            except re.error as e:
                print(f"❌ 정규식 오류: {e}", file=sys.stderr)
                sys.exit(1)
        else:
            self.regex = None

    def should_ignore(self, path: Path) -> bool:
        """경로가 제외 목록에 있는지 확인"""
        for part in path.parts:
            if part in self.ignore_dirs:
                return True
        return False

    def is_binary(self, path: Path) -> bool:
        """바이너리 파일인지 확인"""
        binary_extensions = {
            'exe', 'dll', 'so', 'dylib', 'bin', 'o', 'obj',
            'png', 'jpg', 'jpeg', 'gif', 'bmp', 'zip', 'tar', 'gz',
            'db', 'sqlite', 'iso', 'dmg', 'pdf', 'doc', 'docx'
        }
        return path.suffix.lower().lstrip('.') in binary_extensions

    def match_keyword(self, text: str) -> bool:
        """키워드 매칭"""
        if self.regex:
            return self.regex.search(text) is not None
        else:
            return self.keyword.lower() in text.lower()

    def search_filename(self, path: Path) -> bool:
        """파일명 검색"""
        return self.match_keyword(path.name)

    def search_content(self, path: Path) -> List[Tuple[int, str]]:
        """파일 내용 검색"""
        matches = []

        if self.is_binary(path):
            return matches

        try:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                for line_num, line in enumerate(f, 1):
                    if self.match_keyword(line):
                        matches.append((line_num, line.rstrip()))
        except Exception as e:
            pass  # 읽을 수 없는 파일은 무시

        return matches

    def search_file(self, file_path: Path) -> List[dict]:
        """단일 파일 검색"""
        results = []

        # 파일명 검색
        if self.search_filename(file_path):
            results.append({
                'type': 'filename',
                'path': str(file_path),
                'line_num': None,
                'content': None
            })

        # 파일 내용 검색
        content_matches = self.search_content(file_path)
        for line_num, content in content_matches:
            results.append({
                'type': 'content',
                'path': str(file_path),
                'line_num': line_num,
                'content': content
            })

        return results

    def search(self) -> List[dict]:
        """전체 검색 실행"""
        files_to_search = []

        # 파일 수집
        for root, dirs, files in os.walk(self.path):
            # 제외 폴더 제거
            dirs[:] = [d for d in dirs if d not in self.ignore_dirs]

            for file in files:
                file_path = Path(root) / file
                if not self.should_ignore(file_path):
                    files_to_search.append(file_path)

        if not files_to_search:
            return []

        # 병렬 검색
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {executor.submit(self.search_file, f): f for f in files_to_search}

            for future in as_completed(futures):
                try:
                    self.results.extend(future.result())
                except Exception as e:
                    pass

        return self.results


def print_results(results: List[dict], limit: Optional[int], name_only: bool = False, content_only: bool = False):
    """결과 출력"""
    if not results:
        print("❌ 검색 결과가 없습니다.")
        return

    # 필터링
    if name_only:
        results = [r for r in results if r['type'] == 'filename']
    elif content_only:
        results = [r for r in results if r['type'] == 'content']

    # 결과 제한
    if limit:
        results = results[:limit]

    if not results:
        print("❌ 검색 결과가 없습니다.")
        return

    print("\n" + "="*80)
    print(f"🔍 검색 결과 ({len(results)}개)")
    print("="*80 + "\n")

    current_file = None
    for result in results:
        if result['type'] == 'filename':
            print(f"📄 {result['path']}")
        else:
            # 같은 파일면 줄만 출력
            if current_file != result['path']:
                print(f"\n📄 {result['path']}")
                current_file = result['path']

            # 줄 번호와 내용 출력
            line_content = result['content'][:100] + "..." if len(result['content']) > 100 else result['content']
            print(f"   {result['line_num']:>5}: {line_content}")

    print("\n" + "="*80)
    print(f"✅ 총 {len(results)}개의 결과를 찾았습니다.")
    print("="*80)


def main():
    parser = argparse.ArgumentParser(
        description='빠른 파일 검색 도구 - 파일명과 파일 내용을 동시에 검색합니다.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
사용 예시:
  python fsearch.py "검색어"                              # 현재 디렉토리 검색
  python fsearch.py "검색어" -p C:\\Users               # 특정 경로 검색
  python fsearch.py "검색어" -i ".git,__pycache__"      # 제외 폴더 지정
  python fsearch.py "검색어" --name-only                # 파일명만 검색
  python fsearch.py "검색어" -r                          # 정규식으로 검색
        '''
    )

    parser.add_argument('keyword', help='검색할 키워드')
    parser.add_argument('-p', '--path', type=Path, default=Path('.'), help='검색 경로 (기본값: 현재 디렉토리)')
    parser.add_argument('-i', '--ignore', help='제외할 폴더 (쉼표로 구분, 예: .git,__pycache__,node_modules)')
    parser.add_argument('-n', '--name-only', action='store_true', help='파일명만 검색')
    parser.add_argument('-c', '--content-only', action='store_true', help='파일 내용만 검색')
    parser.add_argument('-l', '--limit', type=int, help='결과 개수 제한')
    parser.add_argument('-r', '--regex', action='store_true', help='정규식 사용')
    parser.add_argument('-w', '--workers', type=int, default=8, help='병렬 처리 스레드 수 (기본값: 8)')

    args = parser.parse_args()

    # 검색 경로 확인
    if not args.path.exists():
        print(f"❌ 경로를 찾을 수 없습니다: {args.path}", file=sys.stderr)
        sys.exit(1)

    # 제외 폴더 파싱
    ignore_dirs = []
    if args.ignore:
        ignore_dirs = [d.strip() for d in args.ignore.split(',')]

    # 기본 제외 폴더
    default_ignores = {'.git', '__pycache__', 'node_modules', '.venv', 'venv', '.idea', '.vscode'}
    ignore_dirs.extend(default_ignores)

    print(f"🔍 검색어: '{args.keyword}'")
    print(f"📁 경로: {args.path.resolve()}")
    print(f"🔧 옵션: 정규식={args.regex}, 스레드={args.workers}")
    print()

    # 검색 실행
    searcher = FileSearcher(
        keyword=args.keyword,
        path=args.path,
        ignore_dirs=ignore_dirs,
        use_regex=args.regex,
        max_workers=args.workers
    )

    results = searcher.search()

    # 결과 출력
    print_results(
        results,
        limit=args.limit,
        name_only=args.name_only,
        content_only=args.content_only
    )


if __name__ == '__main__':
    main()
