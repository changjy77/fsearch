# 🔍 fsearch - 빠른 파일 검색 도구

Windows 개인 PC에서 **파일명과 파일 내용을 동시에** 빠르게 검색하는 도구입니다.

## ⚡ 특징

- ✅ **파일명 검색** - 빠른 파일명 검색
- ✅ **파일 내용 검색** - 텍스트 파일 내용 실시간 검색
- ✅ **멀티스레딩** - 8개 스레드로 병렬 처리 (매우 빠름)
- ✅ **정규식 지원** - 복잡한 패턴 검색 가능
- ✅ **제외 폴더** - .git, node_modules 등 자동 제외
- ✅ **간단한 CLI** - 쉬운 사용법

## 📋 설치

### 요구사항
- Python 3.7+
- Windows 10 이상

### 설정

1. **시스템 PATH에 추가** (선택)
   ```powershell
   # PowerShell (관리자 권한)
   $env:Path += ";D:\클로드\fsearch"
   [System.Environment]::SetEnvironmentVariable("Path", $env:Path, "User")
   ```

2. **아니면 현재 디렉토리에서만 사용**
   ```cmd
   python fsearch.py "검색어"
   ```

## 🚀 사용법

### 기본 사용
```bash
python fsearch.py "검색어"
```

### 특정 경로 검색
```bash
python fsearch.py "검색어" -p C:\Users\username\Documents
```

### 파일명만 검색
```bash
python fsearch.py "검색어" --name-only
```

### 파일 내용만 검색
```bash
python fsearch.py "검색어" --content-only
```

### 정규식 검색
```bash
python fsearch.py "^[a-z]+\.txt$" --regex
```

### 제외 폴더 지정
```bash
python fsearch.py "검색어" --ignore ".git,__pycache__,node_modules"
```

### 결과 개수 제한
```bash
python fsearch.py "검색어" --limit 20
```

### 병렬 처리 스레드 조정
```bash
python fsearch.py "검색어" --workers 4
```

## 📊 실제 사용 예시

### 예 1: 프로젝트에서 함수명 찾기
```bash
python fsearch.py "def main" -p C:\MyProject --content-only
```

### 예 2: 특정 파일명 검색
```bash
python fsearch.py "config" --name-only
```

### 예 3: 오류 메시지 찾기
```bash
python fsearch.py "NullPointerException" -p C:\workspace
```

### 예 4: 정규식으로 복잡한 검색
```bash
python fsearch.py "error\s*:\s*\d+" --regex
```

## ⚙️ 옵션 참고

| 옵션 | 설명 |
|------|------|
| `-p, --path PATH` | 검색 경로 (기본값: 현재 디렉토리) |
| `-i, --ignore FOLDERS` | 제외할 폴더 (쉼표 구분) |
| `-n, --name-only` | 파일명만 검색 |
| `-c, --content-only` | 파일 내용만 검색 |
| `-l, --limit COUNT` | 결과 개수 제한 |
| `-r, --regex` | 정규식 사용 |
| `-w, --workers NUM` | 스레드 수 조정 (기본값: 8) |
| `-h, --help` | 도움말 표시 |

## 🎯 성능

- **작은 폴더** (< 1000개 파일): < 1초
- **중간 폴더** (1000-100000개 파일): 1-10초
- **큰 폴더** (> 100000개 파일): 10-60초

멀티스레딩으로 처리되므로 CPU 코어가 많을수록 빠릅니다.

## 📝 자동 제외 폴더

다음 폴더는 자동으로 제외됩니다:
- `.git`
- `__pycache__`
- `node_modules`
- `.venv`, `venv`
- `.idea`, `.vscode`

`--ignore` 옵션으로 추가 제외 폴더를 지정할 수 있습니다.

## 🐛 바이너리 파일 필터

다음 파일은 자동으로 스킵됩니다 (검색하지 않음):
- 실행파일: `.exe`, `.dll`, `.so`, `.dylib`
- 이미지: `.png`, `.jpg`, `.gif`, `.bmp`
- 아카이브: `.zip`, `.tar`, `.gz`
- 기타: `.db`, `.sqlite`, `.iso`, `.pdf`

## 💡 팁

1. **느린 검색인 경우**
   - 검색 경로를 좀 더 구체적으로 지정
   - `--ignore` 옵션으로 불필요한 폴더 제외
   - `--workers` 수 조정

2. **정확한 검색**
   - 정규식 사용 (`-r` 옵션)
   - 파일명/내용 구분 (`--name-only`, `--content-only`)

3. **대소문자 무시**
   - 기본 검색은 대소문자를 구분하지 않습니다
   - 정규식으로 대소문자 구분 가능: `(?-i:keyword)`

## 🔄 업데이트 및 개선

이 프로젝트는 계속 개선됩니다. 향후 예정 사항:

- [ ] GUI 버전 개발
- [ ] Rust 버전으로 성능 최적화
- [ ] 검색 결과 저장 기능
- [ ] 파일 치환 기능

---

**작성자**: Claude AI  
**라이선스**: MIT  
**최종 업데이트**: 2026-08-01
