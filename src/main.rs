use anyhow::Result;
use clap::Parser;
use colored::Colorize;
use rayon::prelude::*;
use regex::Regex;
use std::fs;
use std::path::PathBuf;
use walkdir::WalkDir;

#[derive(Parser, Debug)]
#[command(name = "fsearch")]
#[command(about = "빠른 파일 검색 도구", long_about = None)]
struct Args {
    /// 검색할 단어
    #[arg(value_name = "KEYWORD")]
    keyword: String,

    /// 검색 경로 (기본값: 현재 디렉토리)
    #[arg(short, long, value_name = "PATH")]
    path: Option<PathBuf>,

    /// 파일명만 검색
    #[arg(short = 'n', long)]
    name_only: bool,

    /// 파일 내용만 검색
    #[arg(short = 'c', long)]
    content_only: bool,

    /// 검색 결과 개수 제한 (기본값: 무제한)
    #[arg(short, long, value_name = "COUNT")]
    limit: Option<usize>,

    /// 제외할 폴더 (쉼표로 구분)
    #[arg(short, long, value_name = "FOLDERS")]
    ignore: Option<String>,

    /// 정규식 사용
    #[arg(short = 'r', long)]
    regex: bool,
}

struct SearchResult {
    path: PathBuf,
    line_number: Option<usize>,
    line_content: Option<String>,
    is_filename_match: bool,
}

fn should_ignore(path: &std::path::Path, ignore_list: &[String]) -> bool {
    for ignore_pattern in ignore_list {
        if let Some(file_name) = path.file_name() {
            if let Some(name_str) = file_name.to_str() {
                if name_str == ignore_pattern {
                    return true;
                }
            }
        }
        if path.to_string_lossy().contains(ignore_pattern) {
            return true;
        }
    }
    false
}

fn search_filename(path: &PathBuf, keyword: &str, use_regex: bool) -> bool {
    if let Some(file_name) = path.file_name() {
        if let Some(name_str) = file_name.to_str() {
            if use_regex {
                if let Ok(re) = Regex::new(keyword) {
                    return re.is_match(name_str);
                }
            } else {
                return name_str.contains(keyword);
            }
        }
    }
    false
}

fn search_content(path: &PathBuf, keyword: &str, use_regex: bool) -> Vec<(usize, String)> {
    let mut results = Vec::new();

    if is_binary(path) {
        return results;
    }

    if let Ok(content) = fs::read_to_string(path) {
        for (line_num, line) in content.lines().enumerate() {
            let matched = if use_regex {
                if let Ok(re) = Regex::new(keyword) {
                    re.is_match(line)
                } else {
                    false
                }
            } else {
                line.contains(keyword)
            };

            if matched {
                results.push((line_num + 1, line.to_string()));
            }
        }
    }

    results
}

fn is_binary(path: &PathBuf) -> bool {
    let binary_extensions = [
        "exe", "dll", "so", "dylib", "bin", "o", "obj",
        "png", "jpg", "jpeg", "gif", "bmp", "zip", "tar", "gz",
        "db", "sqlite", "iso", "dmg",
    ];

    if let Some(ext) = path.extension() {
        if let Some(ext_str) = ext.to_str() {
            return binary_extensions.contains(&ext_str.to_lowercase().as_str());
        }
    }
    false
}

fn main() -> Result<()> {
    let args = Args::parse();

    let search_path = args.path.unwrap_or_else(|| PathBuf::from("."));
    let ignore_list: Vec<String> = args
        .ignore
        .map(|s| s.split(',').map(|x| x.trim().to_string()).collect())
        .unwrap_or_default();

    println!(
        "{}",
        format!("🔍 검색어: '{}'", args.keyword).bold().cyan()
    );
    println!(
        "{}",
        format!("📁 경로: {}", search_path.display()).cyan()
    );
    println!();

    let files: Vec<PathBuf> = WalkDir::new(&search_path)
        .into_iter()
        .filter_map(|e| e.ok())
        .filter(|e| e.file_type().is_file())
        .filter(|e| !should_ignore(e.path(), &ignore_list))
        .map(|e| e.path().to_path_buf())
        .collect();

    if files.is_empty() {
        println!("{}", "검색할 파일이 없습니다.".yellow());
        return Ok(());
    }

    let results: Vec<SearchResult> = files
        .par_iter()
        .flat_map(|file_path| {
            let mut file_results = Vec::new();

            if !args.content_only && search_filename(file_path, &args.keyword, args.regex) {
                file_results.push(SearchResult {
                    path: file_path.clone(),
                    line_number: None,
                    line_content: None,
                    is_filename_match: true,
                });
            }

            if !args.name_only {
                let content_matches = search_content(file_path, &args.keyword, args.regex);
                for (line_num, line_content) in content_matches {
                    file_results.push(SearchResult {
                        path: file_path.clone(),
                        line_number: Some(line_num),
                        line_content: Some(line_content),
                        is_filename_match: false,
                    });
                }
            }

            file_results
        })
        .collect();

    let limited_results: Vec<_> = results
        .into_iter()
        .take(args.limit.unwrap_or(usize::MAX))
        .collect();

    if limited_results.is_empty() {
        println!("{}", "검색 결과가 없습니다.".yellow());
        return Ok(());
    }

    println!("{}:", "검색 결과".bold().green());
    println!();

    for result in &limited_results {
        if result.is_filename_match {
            println!(
                "{} {}",
                "📄".green(),
                result.path.display().to_string().bold()
            );
        } else {
            println!(
                "{}:{} {}",
                result.path.display(),
                result.line_number.unwrap_or(0).to_string().yellow(),
                result.line_content.as_ref().unwrap_or(&String::new()).dimmed()
            );
        }
    }

    println!();
    println!(
        "{}",
        format!(
            "총 {} 개의 결과를 찾았습니다.",
            limited_results.len()
        )
        .bold()
        .green()
    );

    Ok(())
}
