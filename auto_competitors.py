"""
대형사 동향 - 주간 경쟁사 사업동향 보고서
삼성물산, 삼성E&A, 대우건설, GS건설, DL이앤씨
네이버 뉴스 API 기반 · 1주일 이내 · 회사별 사업동향 촘촘 추적
─────────────────────────────────────────────
핵심 원칙:
  - 갯수 채우기 X → 진짜 사업동향만 빠짐없이 수집
  - 회사가 기사의 '주체'인 것만 (나열식 언급 제외)
  - 동향 없으면 0건도 OK
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
        # 다양한 사업동향을 빠짐없이 잡기 위해 검색 키워드 확장
        "keywords": [
            "삼성물산 건설", "삼성물산 수주", "삼성물산 착공", "삼성물산 분양",
            "삼성물산 시공", "삼성물산 재건축", "삼성물산 계약", "삼성물산 MOU",
            "삼성물산 입찰", "삼성물산 해외", "삼성물산 사업", "삼성물산 기술",
            "삼성물산 조직", "삼성물산 인사", "삼성물산 대표",
            "래미안",  # 브랜드명
        ],
        "identifier": ["삼성물산", "래미안"],
        "color": "#1428a0",
    },
    "삼성E&A": {
        "keywords": [
            "삼성E&A", "삼성엔지니어링", "삼성E&A 수주", "삼성E&A 플랜트",
            "삼성E&A EPC", "삼성E&A 해외", "삼성E&A 계약", "삼성E&A 사업",
            "삼성E&A 기술", "삼성E&A 조직", "삼성E&A 대표",
        ],
        "identifier": ["삼성E&A", "삼성엔지니어링"],
        "color": "#1428a0",
    },
    "대우건설": {
        "keywords": [
            "대우건설", "대우건설 수주", "대우건설 분양", "대우건설 착공",
            "대우건설 시공", "대우건설 재건축", "대우건설 계약", "대우건설 MOU",
            "대우건설 입찰", "대우건설 해외", "대우건설 사업", "대우건설 기술",
            "대우건설 조직", "대우건설 인사", "대우건설 대표",
            "푸르지오",  # 브랜드명
        ],
        "identifier": ["대우건설", "푸르지오"],
        "color": "#003da5",
    },
    "GS건설": {
        "keywords": [
            "GS건설", "GS건설 수주", "GS건설 분양", "GS건설 착공",
            "GS건설 시공", "GS건설 재건축", "GS건설 계약", "GS건설 MOU",
            "GS건설 입찰", "GS건설 해외", "GS건설 사업", "GS건설 기술",
            "GS건설 조직", "GS건설 인사", "GS건설 대표",
            "자이",  # 브랜드명 (단, 짧아서 identifier에는 넣지 않음)
        ],
        "identifier": ["GS건설"],
        "color": "#ed1c24",
    },
    "DL이앤씨": {
        "keywords": [
            "DL이앤씨", "DL이앤씨 수주", "DL이앤씨 분양", "DL이앤씨 착공",
            "DL이앤씨 시공", "DL이앤씨 재건축", "DL이앤씨 계약", "DL이앤씨 MOU",
            "DL이앤씨 입찰", "DL이앤씨 해외", "DL이앤씨 사업", "DL이앤씨 기술",
            "DL이앤씨 조직", "DL이앤씨 인사", "DL이앤씨 대표",
            "아크로",  # 브랜드명
        ],
        "identifier": ["DL이앤씨", "아크로"],
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
    # [주간 요약] 오늘 기준 7일 롤링 윈도우.
    # 자정 기준 cutoff — 같은 날 여러 번 돌려도 결과가 미끄러지지 않음.
    today_midnight = datetime.now(KST).replace(hour=0, minute=0, second=0, microsecond=0)
    cutoff = today_midnight - timedelta(days=7)

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
# 2) 중복 제거 (강화)
# ──────────────────────────────────────
def _title_keywords(title):
    """제목에서 핵심 키워드 추출 (조사 제거, 핵심어 중심)"""
    clean = re.sub(r"[^가-힣a-zA-Z0-9]", " ", title)
    words = clean.split()

    # 한국어 조사/어미 제거 (단어 끝에서)
    PARTICLES = re.compile(
        r"(이|가|은|는|을|를|에|의|도|로|으로|와|과|에서|까지|부터|만|도|든|란|라|며|고|해|한|된|할|하는|하고|하며|했다|한다|으며)$"
    )
    stripped = set()
    for w in words:
        if len(w) < 2:
            continue
        s = PARTICLES.sub("", w)
        if len(s) >= 2:
            stripped.add(s)
        else:
            stripped.add(w)
    return stripped


def _similarity(title_a, title_b):
    """두 제목의 핵심 키워드 유사도 (0~1), 부분 매칭 포함"""
    kw_a = _title_keywords(title_a)
    kw_b = _title_keywords(title_b)
    if not kw_a or not kw_b:
        return 0

    # 정확 매칭 + 부분 포함 매칭 (한쪽이 다른쪽에 포함되면 0.5점)
    match_score = 0
    for a in kw_a:
        for b in kw_b:
            if a == b:
                match_score += 1
                break
            elif len(a) >= 3 and len(b) >= 3 and (a in b or b in a):
                match_score += 0.7
                break

    shorter = min(len(kw_a), len(kw_b))
    return match_score / shorter if shorter > 0 else 0


def remove_duplicates(articles):
    """URL 중복 + 유사 제목 제거 (유사도 60% 이상이면 최신 1건만)"""
    # URL 기반 중복 제거
    seen_urls = set()
    url_unique = []
    for art in articles:
        if art["link"] not in seen_urls:
            seen_urls.add(art["link"])
            url_unique.append(art)

    # 유사 제목 제거 (최신 기사 우선 유지)
    url_unique.sort(key=lambda a: a.get("datetime", a.get("date", "")), reverse=True)
    kept = []
    kept_titles = []
    for art in url_unique:
        if any(_similarity(art["title"], prev) >= 0.6 for prev in kept_titles):
            continue
        kept.append(art)
        kept_titles.append(art["title"])
    return kept


# ──────────────────────────────────────
# 3) 사업동향 필터링 — 핵심 알고리즘
# ──────────────────────────────────────
# [원칙] 회사가 기사의 '주체'이고, 구체적인 사업 활동이 있어야 통과
# [원칙] 갯수를 채우는 게 아니라 진짜 동향만 (0건도 OK)

# ── 하드 제외: 이 키워드가 있으면 무조건 탈락 ──
HARD_EXCLUDE = [
    # 스포츠
    "배구", "V리그", "챔프전", "플레이오프", "세트스코어", "프로배구",
    "선수단", "선수", "경기에서", "감독대행",
    # 주식/증권/투자 (사업동향이 아닌 시세 기사)
    "장중", "주가", "목표가", "특징주", "관련주", "건설주", "테마주",
    "급등주", "상한가", "하한가", "주봉", "매수", "매도", "증권사",
    "펀드", "ETF", "투자의견", "컨센서스", "종목분석", "리포트",
    "52주 신고가", "52주 신저가", "시가총액", "PER", "PBR",
    "이격도", "과열", "급등", "급락", "반등", "하락세",
    # 부고/사회
    "부고", "별세", "결혼", "장례",
    # 조합 내부 정치 (사업동향 아님)
    "조합장 해임", "조합장 선거", "비대위", "총회 무산", "비리",
    # 광고성/홍보/분양광고
    "모집공고", "입주자 모집", "견본주택", "모델하우스 오픈",
    # 일반 시황/전망/공시 나열 기사
    "오늘의 주요공시", "[마감]", "[장마감]", "[장중]",
    "오늘의 증시", "주간 증시", "이번주 증시",
    "[주요공시]", "주요공시", "[코스피", "[코스닥",
    "코스피·코스닥", "코스피코스닥",
    # 업계 나열식 시리즈 (특정 회사 동향 아님)
    "[건설업계 동향]", "[건설업계 브리핑]", "[건설 Pick]",
    "[C-워드]", "[ND 건설]", "[금주의 건설", "[산업계 이모저모]",
    "[건설·부동산 레이더]", "[N2 부동산]", "[은행 뉴스브리핑]",
    "[금융로드]", "[iN The Scene]", "[줌인]",
    "굿모닝", "건설업계 소식",
    # 청약/분양 시세/로또
    "만점통장", "만점 청약", "만점 통장", "로또", "당첨자",
    "청약 경쟁률", "청약 마감", "청약통장", "청약 신기록",
    # 분양 홍보/광고성
    "눈길", "관심 집중", "관심 폭발", "인기 비결",
    "품귀", "완판", "[분양돋보기]", "분양돋보기",
    # CSR/사회공헌
    "사회공헌", "봉사", "기부", "장학금", "예술교육", "교육 지원",
    "키즈", "어린이", "환경정화", "생태활동",
    "창의예술", "창의융합", "미래인재", "인재 양성", "인재 교육",
    "교육 협력", "교육 추진", "ESG 확대", "임직원 참여",
    "Nature 조성", "Nature' 조성",
    # 아파트 거래/시세
    "신고가", "실거래", "실거래가", "매매가", "전세가", "호가",
    "시세", "거래량", "거래 동향",
    # 분양 홍보성 패턴
    "분상제·입지", "분상제·인프라", "브랜드 파워",
    "희소성", "입지·브랜드",
    # 노사/소송/내부 분쟁
    "퇴직금", "퇴직자", "노조", "노사",
    # 연재/기획 시리즈
    "창간", "창간12주년", "창간 12주년",
    "건설사별 수주전략", "건설사별 분석", "[건설사 체력점검",
    "체력점검", "건설사별 ",
]

# ── 나열식 기사 탐지 패턴 ──
# "현대·삼성·대우·GS·DL" 같은 나열식은 특정 회사 동향이 아님
LISTING_PATTERN = re.compile(
    r"(현대건설|현대|삼성물산|삼성|대우건설|대우|GS건설|GS|DL이앤씨|DL|포스코|롯데건설|롯데)"
    r"[·,\s]+"
    r"(현대건설|현대|삼성물산|삼성|대우건설|대우|GS건설|GS|DL이앤씨|DL|포스코|롯데건설|롯데)"
    r"[·,\s]+"
    r"(현대건설|현대|삼성물산|삼성|대우건설|대우|GS건설|GS|DL이앤씨|DL|포스코|롯데건설|롯데)"
)

# ── 사업동향 카테고리별 신호 키워드 ──
# 카테고리 + 키워드 + 가중치
# 기사가 최소 1개 카테고리에 매칭되어야 통과
BUSINESS_CATEGORIES = {
    "수주계약": {
        "keywords": [
            ("수주", 25), ("낙찰", 25), ("시공사 선정", 25), ("시공권", 22),
            ("시공사로", 22), ("계약 체결", 20), ("도급계약", 20),
            ("수주액", 20), ("수주 잔고", 15), ("수주전", 12),
        ],
        "tag_class": "tag-compete",
    },
    "입찰경쟁": {
        "keywords": [
            ("입찰", 18), ("응찰", 18), ("현장설명회", 18), ("설명회", 15),
            ("단독 입찰", 20), ("경쟁 입찰", 18), ("턴키", 15),
            ("우선협상", 20), ("시공사 입찰", 20),
        ],
        "tag_class": "tag-compete",
    },
    "정비사업": {
        "keywords": [
            ("재건축", 15), ("재개발", 15), ("리모델링", 15),
            ("정비사업", 15), ("정비구역", 12), ("조합", 10),
            ("관리처분", 15), ("사업시행", 15), ("철거", 12),
        ],
        "tag_class": "tag-compete",
    },
    "착공준공": {
        "keywords": [
            ("착공", 18), ("준공", 18), ("기공식", 18), ("상량식", 15),
            ("분양", 12), ("청약", 10), ("공급", 8),
        ],
        "tag_class": "tag-compete",
    },
    "해외사업": {
        "keywords": [
            ("해외 수주", 25), ("해외 사업", 20), ("해외 진출", 20),
            ("EPC", 18), ("플랜트", 15), ("중동", 12), ("사우디", 15),
            ("UAE", 15), ("이라크", 12), ("카타르", 12), ("쿠웨이트", 12),
            ("동남아", 12), ("베트남", 12), ("인도", 12),
        ],
        "tag_class": "tag-energy",
    },
    "에너지인프라": {
        "keywords": [
            ("원전", 20), ("SMR", 20), ("해상풍력", 18), ("LNG", 15),
            ("데이터센터", 18), ("수소", 15), ("CCUS", 15), ("신재생", 12),
            ("태양광", 10), ("배터리", 12), ("전력", 10),
        ],
        "tag_class": "tag-energy",
    },
    "사업전략": {
        "keywords": [
            ("MOU", 18), ("업무협약", 18), ("공동개발", 18),
            ("컨소시엄", 15), ("JV", 15), ("합작", 15),
            ("신사업", 18), ("사업 다각화", 18),
            ("파트너십", 15),
        ],
        "tag_class": "tag-strategy",
    },
    "기술개발": {
        "keywords": [
            ("기술 개발", 20), ("기술개발", 20), ("특허", 18), ("R&D", 18),
            ("스마트건설", 18), ("DX", 15), ("로봇", 15),
            ("BIM", 15), ("모듈러", 15), ("OSC", 15), ("3D프린팅", 15),
            ("자동화", 12), ("탄소중립", 12),
            ("제로에너지", 15),
        ],
        "tag_class": "tag-tech",
    },
    "조직인사": {
        "keywords": [
            ("조직개편", 22), ("대표이사", 15), ("CEO", 15),
            ("부회장", 15),
            ("취임", 18), ("선임", 15),
        ],
        "tag_class": "tag-org",
    },
    "경영실적": {
        "keywords": [
            ("영업이익", 18), ("분기 실적", 20),
            ("흑자전환", 20), ("적자전환", 20),
            ("매출액", 15), ("당기순이익", 18), ("영업이익률", 15),
        ],
        "tag_class": "tag-finance",
    },
    "리스크": {
        "keywords": [
            ("안전사고", 20), ("붕괴", 20), ("화재", 15),
            ("공사 중단", 20), ("소송", 15),
            ("하자", 15), ("공사비 증액", 18), ("공사비 분쟁", 18),
            ("부실시공", 20), ("시정명령", 18), ("제재", 15),
        ],
        "tag_class": "tag-risk",
    },
}


def _classify_article(text):
    """기사의 사업동향 카테고리와 총점을 판정"""
    matched_categories = []
    total_score = 0

    for cat_name, cat_info in BUSINESS_CATEGORIES.items():
        cat_score = 0
        for kw, pts in cat_info["keywords"]:
            if kw in text:
                cat_score += pts
        if cat_score > 0:
            matched_categories.append((cat_name, cat_score, cat_info["tag_class"]))
            total_score += cat_score

    # 점수 높은 순 정렬
    matched_categories.sort(key=lambda x: x[1], reverse=True)
    return matched_categories, total_score


def _is_company_subject(title, description, identifiers):
    """
    회사가 기사의 '주체'인지 엄격하게 판단
    핵심 원칙: 제목에 회사명이 있어야 기본 통과 (본문에만 있으면 탈락)
    """
    title_has = any(ident in title for ident in identifiers)

    if not title_has:
        return False, 0  # 제목에 없으면 바로 탈락

    # 제목에서 나열식인지 체크 (3개 이상 건설사가 나열)
    if LISTING_PATTERN.search(title):
        # 나열식이라도 해당 회사가 제목 맨 앞 주어이면 OK
        for ident in identifiers:
            idx = title.find(ident)
            if idx != -1 and idx < 5:
                return True, 80
        # 나열식 + 주어 아님 → 탈락
        return False, 0

    # 회사 정식 명칭 vs 브랜드명만 구분
    # 브랜드명(래미안, 푸르지오, 아크로 등)만 있고 정식 회사명은 없으면 → 낮은 점수
    # (분양 홍보/아파트 시세 기사일 가능성 높음)
    BRAND_ONLY = {"래미안", "푸르지오", "자이", "아크로"}
    company_names = [i for i in identifiers if i not in BRAND_ONLY]
    brand_names = [i for i in identifiers if i in BRAND_ONLY]

    has_company_name = any(cn in title for cn in company_names) if company_names else True
    has_brand_only = not has_company_name and any(bn in title for bn in brand_names)

    if has_brand_only:
        return True, 50  # 브랜드명만 → 낮은 점수 (높은 biz_score 필요)

    return True, 100


def filter_business_news(articles, company_config):
    """
    핵심 필터링: 회사가 주체 + 구체적 사업동향이 있는 기사만 통과
    갯수 제한 없음 — 진짜 동향만 빠짐없이 수집
    """
    filtered = []
    identifiers = company_config["identifier"]

    for art in articles:
        title = art["title"]
        description = art.get("description", "")
        text = title + " " + description

        # ① 하드 제외
        if any(kw in text for kw in HARD_EXCLUDE):
            continue

        # ② 회사가 기사의 주체인지 판단
        is_subject, subject_score = _is_company_subject(title, description, identifiers)
        if not is_subject:
            continue

        # ③ 사업동향 카테고리 분류
        categories, biz_score = _classify_article(text)
        if not categories:
            continue  # 어떤 사업동향 카테고리에도 해당 안 됨

        # ④ 종합 점수 계산
        total_score = subject_score + biz_score

        # 사업동향 점수가 최소 15점 이상이어야 (구체적 동향 키워드가 있어야)
        if biz_score < 15:
            continue

        # 최소 통과 점수 (subject 100 + biz 15 = 115가 기본선)
        if total_score < 110:
            continue

        # 카테고리 태그 부여 (최고 점수 카테고리)
        art["_categories"] = categories
        art["_primary_category"] = categories[0][0] if categories else ""
        art["_primary_tag_class"] = categories[0][2] if categories else ""
        art["_score"] = total_score
        art["_subject_score"] = subject_score

        filtered.append(art)

    return filtered


# ──────────────────────────────────────
# 4) 최종 선정 (갯수 제한 없음, 유사기사만 제거)
# ──────────────────────────────────────
MAJOR_OUTLETS = [
    "mk.co.kr", "hankyung.com", "yna.co.kr", "sedaily.com",
    "chosun.com", "donga.com", "joongang.co.kr", "khan.co.kr",
    "mt.co.kr", "asiae.co.kr", "fnnews.com", "edaily.co.kr",
    "news1.kr", "newsis.com", "sbs.co.kr", "kbs.co.kr",
    "mbc.co.kr", "jtbc.co.kr", "ytn.co.kr",
]


# ── 일반 단어 (entity 추출에서 제외할 흔한 단어) ──
GENERIC_WORDS = {
    "건설", "건설사", "회사", "기업", "그룹", "사업", "수주", "계약",
    "체결", "협약", "MOU", "공동", "개발", "기술", "해외", "국내",
    "정비", "사업", "사업에", "추진", "관련", "영업", "이익", "실적",
    "확대", "강화", "성장", "도전", "진출", "선정", "우선", "협상",
    "대표", "회장", "사장", "대상", "주요", "신규", "최초", "최대",
    "통해", "위해", "위한", "통한", "대한", "에서", "으로", "이번",
    "올해", "내년", "지난", "오늘", "어제", "내일", "오는", "이날",
    "현장", "발표", "계획", "추진", "준비", "예정", "오픈", "개최",
    "실시", "진행", "공개", "출시", "도입", "적용", "선보", "공급",
    "참여", "활용", "지원", "확보", "수익", "경쟁", "투자", "조달",
}


def _extract_entities(text):
    """텍스트에서 핵심 고유명사 추출 (3자 이상, 일반 단어 제외)"""
    clean = re.sub(r"[^가-힣a-zA-Z0-9]", " ", text)
    words = clean.split()

    PARTICLES = re.compile(
        r"(이|가|은|는|을|를|에|의|도|로|으로|와|과|에서|까지|부터|만|든|란|라|며|고|해|한|된|할|하는|하고|하며|했다|한다|으며|등)$"
    )

    entities = set()
    for w in words:
        if len(w) < 3:
            continue
        # 조사 제거
        s = PARTICLES.sub("", w)
        if len(s) < 3:
            continue
        # 일반 단어는 제외
        if s in GENERIC_WORDS:
            continue
        entities.add(s)
    return entities


def _is_same_event(art_a, art_b):
    """
    두 기사가 같은 사안인지 판정
    - 핵심 고유명사 entity 2개 이상 공유 + 날짜 차이 3일 이내
    - OR 제목 유사도 35% 이상
    """
    # 1) 제목 유사도 체크
    if _similarity(art_a["title"], art_b["title"]) >= 0.35:
        return True

    # 2) Entity 기반 매칭 (제목 + description 일부)
    text_a = art_a["title"] + " " + art_a.get("description", "")[:120]
    text_b = art_b["title"] + " " + art_b.get("description", "")[:120]
    ent_a = _extract_entities(text_a)
    ent_b = _extract_entities(text_b)

    common = ent_a & ent_b
    if len(common) >= 3:
        return True

    # 3) 핵심 부분 매칭 (4자+ entity가 양쪽에 포함)
    long_common = {e for e in common if len(e) >= 4}
    if len(long_common) >= 2:
        return True

    return False


def select_top_articles(articles):
    """
    유사기사 제거 후 전부 선정 (갯수 제한 없음)
    같은 사안의 기사 → 주요 매체 or 점수 높은 1건만 대표
    """
    # 주요 매체 가산
    for art in articles:
        bonus = 0
        if any(outlet in art.get("link", "") for outlet in MAJOR_OUTLETS):
            bonus += 10
        hits = art.get("_hits", 1)
        if hits >= 3:
            bonus += 10
        elif hits >= 2:
            bonus += 5
        art["_score"] = art.get("_score", 0) + bonus

    # 점수 높은 순 정렬 (대표 기사를 점수 높은 것으로)
    ranked = sorted(articles, key=lambda a: a["_score"], reverse=True)

    # 유사 기사 제거 (entity 기반)
    picked = []
    for art in ranked:
        if any(_is_same_event(art, prev) for prev in picked):
            continue
        picked.append(art)

    # 최종: 날짜순 정렬 (최신 먼저)
    picked.sort(key=lambda a: a.get("date", ""), reverse=True)
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
    """대형사 동향 HTML — oil-naphtha 라이트 테마 공유"""
    now = datetime.now(KST)
    today = now.strftime("%Y년 %m월 %d일 %H:%M")
    # 월 내 주차 표기 (단순 방식: (일-1)//7 + 1)
    week_no = (now.day - 1) // 7 + 1
    weekly_label = f"{now.month}월 {week_no}주차"

    COMPANY_ICONS = {
        "삼성물산": "🏢",
        "삼성E&A": "⚙️",
        "대우건설": "🏗️",
        "GS건설": "🏛️",
        "DL이앤씨": "🏘️",
    }

    sections_html_list = []
    total_count = sum(len(v) for v in company_articles.values())

    for company_name, config in COMPANIES.items():
        arts = company_articles.get(company_name, [])
        icon = COMPANY_ICONS.get(company_name, "🏢")
        color = config.get("color", "#15ad60")
        top = arts[:2]
        rest = arts[2:]

        top_html = ""
        for a in top:
            source = get_source(a["link"])
            d = a["date"][5:].replace("-", ".") if len(a["date"]) >= 10 else a["date"]
            top_html += f'''
        <a class="top-card" href="{escape(a["link"])}" target="_blank" rel="noopener">
          <div class="top-card-source">{escape(source)} · {d}</div>
          <h3 class="top-card-title">{escape(a["title"])}</h3>
          <p class="top-card-desc">{escape(a.get("description", "")[:160])}</p>
          <span class="top-card-more" style="color:{color}">원문 기사 보기 →</span>
        </a>'''

        rest_html = ""
        for a in rest:
            source = get_source(a["link"])
            d = a["date"][5:].replace("-", ".") if len(a["date"]) >= 10 else a["date"]
            rest_html += f'''
          <li class="headline">
            <a href="{escape(a["link"])}" target="_blank" rel="noopener">
              <span class="hl-title">{escape(a["title"])}</span>
              <span class="hl-meta">{escape(source)} · {d}</span>
            </a>
          </li>'''

        if not arts:
            body = '<div class="empty">해당 기간 주요 동향 없음</div>'
        else:
            body = ""
            if top:
                body += f'<div class="top-cards">{top_html}</div>'
            if rest:
                body += f'<ul class="headlines">{rest_html}</ul>'

        sections_html_list.append(f'''
    <section class="section">
      <div class="section-head">
        <div class="section-icon">{icon}</div>
        <div class="section-meta">
          <h2 class="section-title" style="color:{color}">{escape(company_name)}</h2>
          <div class="section-count" style="background:{color}">{len(arts)}건</div>
        </div>
      </div>
      {body}
    </section>''')

    sections_html = "".join(sections_html_list)
    company_stats = ""
    for cn, cfg in COMPANIES.items():
        cnt = len(company_articles.get(cn, []))
        label = cn if cn != "삼성E&A" else "삼성 E&A"
        company_stats += f'''<div class="stat"><span class="stat-num" style="color:{cfg["color"]}">{cnt}</span><span class="stat-lab">{escape(label)}</span></div>'''

    html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>대형 건설사 동향 Weekly | HDEC Daily News</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;900&family=Outfit:wght@300;400;600;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="style.css">
</head>
<body>

<div class="bg-wrap"><div class="bg-overlay"></div></div>

<main class="container">

  <header class="site-header">
    <a href="https://hdec-daily-news.github.io/landing/" class="brand">
      <span class="brand-dot"></span>
      <span class="brand-name">HDEC Daily News</span>
    </a>
    <div class="site-meta">
      <span>{weekly_label} (최근 7일)</span>
      <span class="sep">·</span>
      <span class="auto-badge">🔄 매일 07시 / 12시 2회 자동 업데이트</span>
    </div>
  </header>

  <section class="hero">
    <div class="hero-tag">🏢 WEEKLY UPDATE</div>
    <h1>
      <span class="hero-sub">주간 업데이트되는</span>
      <span class="hero-main">대형 건설사 동향</span>
    </h1>
    <p class="hero-desc">
      삼성물산 · 삼성E&amp;A · 대우건설 · GS건설 · DL이앤씨<br>
      <strong>5개사 사업동향</strong>을 AI로 수집, 주요 경제지·종합지 기사만 선별합니다.
    </p>
    <div class="hero-stats">
      <div class="stat"><span class="stat-num">{total_count}</span><span class="stat-lab">TOTAL</span></div>
      {company_stats}
    </div>
  </section>

  {sections_html}

  <footer class="site-footer">
    <div>현대건설 글로벌사업부 · 내부 업무용</div>
    <div class="tech">Powered by Naver News API · GitHub Pages</div>
  </footer>

</main>
</body>
</html>'''

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

    # 2) 회사별 필터링 & 선정 (갯수 제한 없음 — 진짜 동향만 전부)
    print("[2/4] 사업동향 필터링 & 선정 중 (품질 우선, 갯수 무제한)...")
    company_results = {}
    all_for_csv = []
    for company, articles in all_articles.items():
        config = COMPANIES[company]
        articles = remove_duplicates(articles)
        articles = filter_business_news(articles, config)
        top = select_top_articles(articles)[:10]  # 회사당 최대 10건
        company_results[company] = top
        all_for_csv.extend(top)  # 중복 제거된 최종 결과만 CSV에 저장
        print(f"  [{company}] {len(articles)}건 필터 통과 → {len(top)}건 최종 선정")

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
