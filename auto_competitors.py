"""
대형사 동향 - 주간 경쟁사 사업동향 보고서
삼성물산, 삼성E&A, 대우건설, GS건설, DL이앤씨
네이버 뉴스 API 기반 · 1주일 이내 · 회사별 사업/수주 동향 중심
"""

import csv
import os
import re
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
from datetime import datetime, timedelta, timezone
from html import unescape

KST = timezone(timedelta(hours=9))

# ──────────────────────────────────────
# 설정
# ──────────────────────────────────────
COMPANIES = {
    "삼성물산": {
        "keywords": ["삼성물산 건설", "삼성물산 수주", "삼성물산 착공", "삼성물산 분양",
                      "삼성물산 시공", "삼성물산 재건축"],
        "identifier": ["삼성물산"],  # 제목/본문에 이 단어가 있어야 해당 회사 기사로 인정
        "color": "#1428a0",
    },
    "삼성E&A": {
        "keywords": ["삼성E&A", "삼성엔지니어링", "삼성E&A 수주", "삼성E&A 플랜트",
                      "삼성E&A EPC", "삼성E&A 해외"],
        "identifier": ["삼성E&A", "삼성엔지니어링"],
        "color": "#1428a0",
    },
    "대우건설": {
        "keywords": ["대우건설", "대우건설 수주", "대우건설 분양", "대우건설 착공",
                      "대우건설 시공", "대우건설 재건축"],
        "identifier": ["대우건설"],
        "color": "#003da5",
    },
    "GS건설": {
        "keywords": ["GS건설", "GS건설 수주", "GS건설 분양", "GS건설 착공",
                      "GS건설 시공", "GS건설 재건축"],
        "identifier": ["GS건설"],
        "color": "#ed1c24",
    },
    "DL이앤씨": {
        "keywords": ["DL이앤씨", "DL이앤씨 수주", "DL이앤씨 분양", "DL이앤씨 착공",
                      "DL이앤씨 시공", "DL이앤씨 재건축"],
        "identifier": ["DL이앤씨"],
        "color": "#e95c0c",
    },
}

NAVER_CLIENT_ID = "4EpC74MmQmbBp2bpWpI5"
NAVER_CLIENT_SECRET = "uxqj17VklI"

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_PATH = os.path.join(OUTPUT_DIR, "index.html")
CSV_PATH = os.path.join(OUTPUT_DIR, "대형사_동향.csv")

# ──────────────────────────────────────
# 1) 네이버 뉴스 API 수집 (1주일)
# ──────────────────────────────────────
def collect_news():
    url = "https://openapi.naver.com/v1/search/news.json"
    headers = {
        "X-Naver-Client-Id": NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET,
    }
    cutoff = datetime.now(KST) - timedelta(days=7)

    all_articles = {}

    for company, config in COMPANIES.items():
        seen_links = {}
        hit_count = {}

        for keyword in config["keywords"]:
            print(f"    [{company}] 검색: '{keyword}'")
            kw_count = 0
            for start in range(1, 400, 100):
                params = {"query": keyword, "display": 100, "start": start, "sort": "date"}
                try:
                    resp = requests.get(url, headers=headers, params=params, timeout=10, verify=False)
                    resp.raise_for_status()
                    data = resp.json()
                except Exception as e:
                    print(f"    [API 오류] {e}")
                    break

                items = data.get("items", [])
                if not items:
                    break

                stop = False
                for item in items:
                    try:
                        pub_date = datetime.strptime(
                            item["pubDate"], "%a, %d %b %Y %H:%M:%S %z"
                        ).astimezone(KST)
                    except ValueError:
                        continue

                    if pub_date < cutoff:
                        stop = True
                        break

                    title = re.sub(r"<.*?>", "", unescape(item.get("title", "")))
                    link = item.get("originallink") or item.get("link", "")
                    description = re.sub(r"<.*?>", "", unescape(item.get("description", "")))

                    if link not in seen_links:
                        seen_links[link] = {
                            "date": pub_date.strftime("%Y-%m-%d"),
                            "datetime": pub_date,
                            "title": title,
                            "link": link,
                            "description": description,
                            "company": company,
                        }
                        hit_count[link] = 0
                        kw_count += 1
                    hit_count[link] += 1

                if stop:
                    break
            print(f"    → {kw_count}건 신규")

        articles = []
        for link, art in seen_links.items():
            art["_hits"] = hit_count[link]
            articles.append(art)

        all_articles[company] = articles
        print(f"  [{company}] 총 {len(articles)}건\n")

    return all_articles


# ──────────────────────────────────────
# 2) 중복 제거
# ──────────────────────────────────────
def remove_duplicates(articles):
    seen = set()
    unique = []
    for art in articles:
        clean = re.sub(r"[^가-힣a-zA-Z0-9]", "", art["title"])
        key = clean[:20]
        if key not in seen:
            seen.add(key)
            unique.append(art)
    return unique


# ──────────────────────────────────────
# 3) 사업동향 관련 기사만 필터링 (핵심)
# ──────────────────────────────────────
EXCLUDE_KEYWORDS = [
    # 스포츠
    "배구", "V리그", "챔프전", "플레이오프", "세트스코어", "프로배구",
    "감독", "선수", "경기에서",
    # 주식/주가/증권
    "장중", "주가", "목표가", "특징주", "관련주", "건설주", "테마주",
    "급등주", "상한가", "하한가", "주봉", "매수", "매도", "증권사",
    "펀드", "ETF", "리포트", "투자의견", "컨센서스",
    # 인사/부고/사회
    "부고", "별세", "결혼",
    # 조합 내부 분쟁
    "조합장 해임", "조합장 선거", "비대위", "총회 무산",
]

# 사업동향/수주동향 신호 키워드 — 이 중 하나라도 있어야 통과
BUSINESS_SIGNALS = [
    # 수주/입찰
    "수주", "입찰", "응찰", "낙찰", "시공사 선정", "시공사", "시공권", "수주전",
    "재건축", "재개발", "리모델링", "정비사업", "도급", "턴키", "책임준공",
    "계약", "발주", "낙찰",
    # 착공/준공/분양
    "착공", "준공", "분양", "공급", "청약",
    # 해외/플랜트/에너지
    "해외", "EPC", "플랜트", "원전", "해상풍력", "LNG", "데이터센터",
    "수소", "SMR", "CCUS", "신재생",
    # 사업 확장/전략
    "사업", "신사업", "진출", "확장", "전략", "MOU", "협약", "공동개발",
    "파트너", "컨소시엄", "JV",
    # 경영/실적
    "매출", "영업이익", "실적", "흑자", "적자", "수익성",
    "CEO", "대표이사", "사장",
    # 리스크/이슈
    "사고", "안전", "지연", "중단", "소송", "분쟁", "하자",
    "공사비", "원가", "비용",
]


def filter_business_news(articles, company_config):
    """해당 회사가 주체인 사업동향 기사만 필터링"""
    filtered = []
    identifiers = company_config["identifier"]

    for art in articles:
        text = art["title"] + " " + art.get("description", "")
        title = art["title"]

        # 제외 키워드 체크
        if any(kw in text for kw in EXCLUDE_KEYWORDS):
            continue

        # 회사명이 제목이나 본문에 있어야 함
        if not any(ident in text for ident in identifiers):
            continue

        # 사업동향 관련 키워드가 있어야 함
        if not any(kw in text for kw in BUSINESS_SIGNALS):
            continue

        filtered.append(art)
    return filtered


# ──────────────────────────────────────
# 4) 스코어링 & 상위 기사 선정
# ──────────────────────────────────────
SCORE_KEYWORDS = [
    # 수주 (가장 높은 가중치)
    ("수주", 20), ("낙찰", 20), ("시공사 선정", 20), ("시공권", 18),
    ("입찰", 15), ("응찰", 15), ("수주전", 15),
    # 정비사업
    ("재건축", 12), ("재개발", 12), ("리모델링", 12), ("정비사업", 12),
    # 착공/분양
    ("착공", 12), ("준공", 10), ("분양", 10), ("공급", 8), ("청약", 8),
    # 해외/플랜트
    ("해외", 15), ("EPC", 15), ("플랜트", 12), ("원전", 18),
    ("해상풍력", 15), ("LNG", 12), ("데이터센터", 12),
    # 경영
    ("매출", 10), ("영업이익", 12), ("실적", 10),
    ("CEO", 8), ("대표이사", 8),
    # 전략
    ("단독", 15), ("첫", 12), ("최초", 15), ("MOU", 10), ("협약", 10),
    ("컨소시엄", 12), ("JV", 12), ("신사업", 12),
    # 리스크
    ("사고", 12), ("지연", 10), ("중단", 12), ("소송", 12),
]

MAJOR_OUTLETS = [
    "mk.co.kr", "hankyung.com", "yna.co.kr", "sedaily.com",
    "chosun.com", "donga.com", "joongang.co.kr", "khan.co.kr",
    "mt.co.kr", "asiae.co.kr", "fnnews.com", "edaily.co.kr",
    "news1.kr", "newsis.com", "sbs.co.kr", "kbs.co.kr",
    "mbc.co.kr", "jtbc.co.kr", "ytn.co.kr",
]


def score_article(art):
    text = art["title"] + " " + art.get("description", "")
    score = 0
    for kw, pts in SCORE_KEYWORDS:
        if kw in text:
            score += pts
    # 화제성 가산
    hits = art.get("_hits", 1)
    if hits >= 3:
        score += 15
    elif hits >= 2:
        score += 8
    # 주요 매체 가산
    if any(outlet in art.get("link", "") for outlet in MAJOR_OUTLETS):
        score += 6
    # 제목에 회사명 직접 언급 가산
    company = art.get("company", "")
    config = COMPANIES.get(company, {})
    if any(ident in art["title"] for ident in config.get("identifier", [])):
        score += 10
    return score


def _title_keywords(title):
    clean = re.sub(r"[^가-힣a-zA-Z0-9]", " ", title)
    return set(w for w in clean.split() if len(w) >= 2)


def _is_similar(title, seen_titles):
    kw_new = _title_keywords(title)
    if not kw_new:
        return False
    for prev_title in seen_titles:
        kw_prev = _title_keywords(prev_title)
        if not kw_prev:
            continue
        overlap = len(kw_new & kw_prev)
        shorter = min(len(kw_new), len(kw_prev))
        if shorter > 0 and overlap / shorter >= 0.5:
            return True
    return False


def select_top_articles(articles, max_count=5):
    """상위 기사 선정 (유사기사 제거)"""
    for art in articles:
        art["_score"] = score_article(art)
    ranked = sorted(articles, key=lambda a: a["_score"], reverse=True)

    picked = []
    seen_titles = []
    for art in ranked:
        if len(picked) >= max_count:
            break
        if _is_similar(art["title"], seen_titles):
            continue
        picked.append(art)
        seen_titles.append(art["title"])
    return picked


# ──────────────────────────────────────
# 5) HTML 보고서 생성 (한 장 보고서 스타일)
# ──────────────────────────────────────
def get_source(link):
    import urllib.parse
    domain = urllib.parse.urlparse(link).netloc.replace("www.", "")
    source_map = {
        "mk.co.kr": "매일경제", "hankyung.com": "한국경제",
        "yna.co.kr": "연합뉴스", "sedaily.com": "서울경제",
        "dnews.co.kr": "대한경제", "chosun.com": "조선일보",
        "khan.co.kr": "경향신문", "mt.co.kr": "머니투데이",
        "asiae.co.kr": "아시아경제", "view.asiae.co.kr": "아시아경제",
        "fnnews.com": "파이낸셜뉴스", "edaily.co.kr": "이데일리",
        "dt.co.kr": "디지털타임스", "news1.kr": "뉴스1",
        "newsis.com": "뉴시스", "ajunews.com": "아주경제",
        "sbs.co.kr": "SBS", "biz.sbs.co.kr": "SBS비즈",
        "jtbc.co.kr": "JTBC", "news.jtbc.co.kr": "JTBC",
        "kbs.co.kr": "KBS", "mbc.co.kr": "MBC",
        "ytn.co.kr": "YTN", "donga.com": "동아일보",
        "joongang.co.kr": "중앙일보", "hani.co.kr": "한겨레",
        "theguru.co.kr": "더구루", "newspim.com": "뉴스핌",
        "etoday.co.kr": "이투데이", "thebell.co.kr": "더벨",
        "dealsite.co.kr": "딜사이트", "investchosun.com": "인베스트조선",
        "ceoscoredaily.com": "CEO스코어데일리",
    }
    return source_map.get(domain, domain)


def escape(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def generate_html(company_articles):
    now = datetime.now(KST)
    today = now.strftime("%Y년 %m월 %d일")
    today_short = now.strftime("%Y.%m.%d")
    week_ago = (now - timedelta(days=7)).strftime("%m.%d")
    period = f"{week_ago} ~ {now.strftime('%m.%d')}"

    # 회사별 섹션 HTML
    company_sections = ""
    total_count = 0

    for company_name, config in COMPANIES.items():
        arts = company_articles.get(company_name, [])
        total_count += len(arts)

        # 기사 리스트 생성
        if not arts:
            articles_html = '<div class="no-news">해당 기간 주요 동향 없음</div>'
        else:
            articles_html = ""
            for art in arts:
                source = get_source(art["link"])
                articles_html += f"""
                <div class="news-item">
                    <a href="{escape(art['link'])}" target="_blank" class="news-title">{escape(art['title'])}</a>
                    <span class="news-meta">{source} · {art['date']}</span>
                </div>"""

        company_sections += f"""
        <div class="company-block">
            <div class="company-label" style="border-left-color:{config['color']}">
                <span class="company-name">{company_name}</span>
                <span class="company-count">{len(arts)}건</span>
            </div>
            <div class="news-list">{articles_html}
            </div>
        </div>"""

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>대형사 동향 - {today}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700;800&display=swap');
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Noto Sans KR', sans-serif; background: #f0f2f5; color: #222; }}

        /* 보고서 용지 */
        .report {{ max-width: 900px; margin: 24px auto; background: #fff; box-shadow: 0 1px 8px rgba(0,0,0,0.08); }}

        /* 헤더 */
        .report-header {{ padding: 32px 40px 24px; border-bottom: 3px solid #1a1a1a; }}
        .report-header h1 {{ font-size: 26px; font-weight: 800; color: #1a1a1a; letter-spacing: 2px; }}
        .report-meta {{ display: flex; justify-content: space-between; align-items: center; margin-top: 8px; }}
        .report-period {{ font-size: 14px; color: #666; font-weight: 500; }}
        .report-date {{ font-size: 13px; color: #999; }}

        /* 요약 바 */
        .summary {{ display: flex; gap: 0; padding: 0 40px; border-bottom: 1px solid #e8e8e8; }}
        .summary-item {{ flex: 1; padding: 16px 0; text-align: center; border-right: 1px solid #e8e8e8; }}
        .summary-item:last-child {{ border-right: none; }}
        .summary-company {{ font-size: 13px; font-weight: 700; color: #333; }}
        .summary-num {{ font-size: 22px; font-weight: 800; margin-top: 4px; }}

        /* 회사별 블록 */
        .report-body {{ padding: 24px 40px 32px; }}

        .company-block {{ margin-bottom: 24px; }}
        .company-block:last-child {{ margin-bottom: 0; }}

        .company-label {{ display: flex; align-items: center; gap: 10px; padding: 8px 0 8px 14px; border-left: 4px solid #333; margin-bottom: 8px; }}
        .company-name {{ font-size: 16px; font-weight: 700; color: #1a1a1a; }}
        .company-count {{ font-size: 12px; color: #999; }}

        .news-list {{ padding-left: 18px; }}

        .news-item {{ padding: 6px 0; display: flex; flex-direction: column; gap: 2px; }}
        .news-item + .news-item {{ border-top: 1px solid #f5f5f5; }}

        .news-title {{ font-size: 14px; font-weight: 500; color: #1a1a1a; text-decoration: none; line-height: 1.6; }}
        .news-title:hover {{ color: #2c3e50; text-decoration: underline; }}

        .news-meta {{ font-size: 11px; color: #aaa; }}

        .no-news {{ font-size: 13px; color: #bbb; padding: 8px 0; font-style: italic; }}

        /* 푸터 */
        .report-footer {{ padding: 16px 40px; border-top: 1px solid #e8e8e8; text-align: center; font-size: 11px; color: #bbb; }}

        /* 인쇄 최적화 */
        @media print {{
            body {{ background: #fff; }}
            .report {{ box-shadow: none; margin: 0; }}
            .report-header {{ padding: 20px 30px 16px; }}
            .report-body {{ padding: 16px 30px 20px; }}
            .summary {{ padding: 0 30px; }}
            .news-title {{ color: #000 !important; }}
            .news-title:hover {{ text-decoration: none; }}
        }}

        @media (max-width: 640px) {{
            .report {{ margin: 0; }}
            .report-header, .report-body {{ padding-left: 20px; padding-right: 20px; }}
            .summary {{ padding: 0 20px; flex-wrap: wrap; }}
            .summary-item {{ flex-basis: 33%; }}
            .report-header h1 {{ font-size: 20px; }}
        }}
    </style>
</head>
<body>
    <div class="report">
        <div class="report-header">
            <h1>대형사 동향</h1>
            <div class="report-meta">
                <span class="report-period">{period} (주간)</span>
                <span class="report-date">작성일 {today}</span>
            </div>
        </div>

        <div class="summary">
"""

    for company_name, config in COMPANIES.items():
        count = len(company_articles.get(company_name, []))
        html += f"""            <div class="summary-item">
                <div class="summary-company">{company_name}</div>
                <div class="summary-num" style="color:{config['color']}">{count}</div>
            </div>
"""

    html += f"""        </div>

        <div class="report-body">
{company_sections}
        </div>

        <div class="report-footer">
            대형사 동향 · 네이버 뉴스 API 기반 사업동향 자동 수집 · {today_short}
        </div>
    </div>
</body>
</html>"""

    return html


# ──────────────────────────────────────
# 메인
# ──────────────────────────────────────
def main():
    print(f"\n{'='*50}")
    print(f"  대형사 동향 - 주간 경쟁사 사업동향 보고서")
    print(f"  {datetime.now(KST).strftime('%Y-%m-%d %H:%M')}")
    print(f"{'='*50}\n")

    # 1) 수집
    print("[1/4] 대형사 뉴스 수집 중 (최근 1주일)...")
    all_articles = collect_news()

    # 2) 회사별 필터링 & 선정
    print("[2/4] 사업동향 필터링 & 상위 기사 선정 중...")
    company_results = {}
    all_for_csv = []
    for company, articles in all_articles.items():
        config = COMPANIES[company]
        articles = remove_duplicates(articles)
        articles = filter_business_news(articles, config)
        top = select_top_articles(articles, max_count=5)
        company_results[company] = top
        all_for_csv.extend(articles)
        print(f"  [{company}] {len(articles)}건 필터 통과 → 상위 {len(top)}건 선정")

    # 3) HTML 보고서 생성
    print("[3/4] HTML 보고서 생성 중...")
    html = generate_html(company_results)
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  → 저장 완료: {HTML_PATH}")

    # 4) CSV 저장
    print("[4/4] CSV 저장 중...")
    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["company", "date", "title", "link", "description"])
        writer.writeheader()
        for art in sorted(all_for_csv, key=lambda a: (a["company"], a["date"]), reverse=True):
            writer.writerow({
                "company": art.get("company", ""),
                "date": art.get("date", ""),
                "title": art.get("title", ""),
                "link": art.get("link", ""),
                "description": art.get("description", ""),
            })
    print(f"  → 저장 완료: {CSV_PATH}")

    print(f"\n  완료!")


if __name__ == "__main__":
    main()
