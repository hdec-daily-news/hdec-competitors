"""
대형사 동향 - 주간 PPT 보고서 생성
auto_competitors.py가 수집한 CSV 데이터를 기반으로
회사별 사업동향 분석 1장 보고서 PPT 생성
"""

import csv
import os
from datetime import datetime, timedelta, timezone
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

KST = timezone(timedelta(hours=9))
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(OUTPUT_DIR, "대형사_동향.csv")
PPT_PATH = os.path.join(OUTPUT_DIR, "대형사_동향.pptx")

COMPANIES = {
    "삼성물산": {"color": RGBColor(0x14, 0x28, 0xA0)},
    "삼성E&A": {"color": RGBColor(0x14, 0x28, 0xA0)},
    "대우건설": {"color": RGBColor(0x00, 0x3D, 0xA5)},
    "GS건설": {"color": RGBColor(0xED, 0x1C, 0x24)},
    "DL이앤씨": {"color": RGBColor(0xE9, 0x5C, 0x0C)},
}

# ──────────────────────────────────────
# CSV 읽기
# ──────────────────────────────────────
def load_articles():
    """CSV에서 회사별 기사 로드"""
    company_articles = {name: [] for name in COMPANIES}
    if not os.path.exists(CSV_PATH):
        print(f"  [경고] CSV 파일 없음: {CSV_PATH}")
        return company_articles

    with open(CSV_PATH, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            company = row.get("company", "")
            if company in company_articles:
                company_articles[company].append(row)

    return company_articles


# ──────────────────────────────────────
# PPT 생성
# ──────────────────────────────────────
def create_report(company_articles):
    now = datetime.now(KST)
    week_num = (now.day - 1) // 7 + 1
    title_text = f"{now.strftime('%y')}년 {now.month}월 {week_num}주차 대형사 동향"
    week_ago = (now - timedelta(days=7)).strftime("%m.%d")
    period = f"{week_ago} ~ {now.strftime('%m.%d')}"

    prs = Presentation()
    # A4 가로 (레터 대신 와이드에 가까운 비율)
    prs.slide_width = Emu(12192000)   # 13.33 inch → 표준 와이드
    prs.slide_height = Emu(6858000)   # 7.5 inch

    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # ── 배경 ──
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # ── 상단 헤더 바 ──
    header_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Emu(0), Emu(0), prs.slide_width, Emu(680000)
    )
    header_bar.fill.solid()
    header_bar.fill.fore_color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
    header_bar.line.fill.background()

    # 헤더 제목
    txBox = slide.shapes.add_textbox(Emu(400000), Emu(120000), Emu(8000000), Emu(450000))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.font.name = "맑은 고딕"

    # 헤더 기간
    txBox2 = slide.shapes.add_textbox(Emu(9000000), Emu(160000), Emu(3000000), Emu(350000))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = period
    p2.font.size = Pt(14)
    p2.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
    p2.font.name = "맑은 고딕"
    p2.alignment = PP_ALIGN.RIGHT

    # ── 회사별 블록 (테이블 스타일) ──
    y_offset = Emu(800000)
    left_margin = Emu(400000)
    block_width = Emu(11400000)
    company_label_width = Emu(1500000)
    content_width = Emu(9900000)

    for company_name, config in COMPANIES.items():
        arts = company_articles.get(company_name, [])
        color = config["color"]

        # 회사명 컬러 바
        accent_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left_margin, y_offset, Emu(60000), Emu(900000)
        )
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = color
        accent_bar.line.fill.background()

        # 회사명 텍스트
        name_box = slide.shapes.add_textbox(
            left_margin + Emu(100000), y_offset + Emu(20000),
            company_label_width, Emu(300000)
        )
        ntf = name_box.text_frame
        ntf.word_wrap = True
        np = ntf.paragraphs[0]
        np.text = company_name
        np.font.size = Pt(13)
        np.font.bold = True
        np.font.color.rgb = color
        np.font.name = "맑은 고딕"

        # 기사 건수
        count_box = slide.shapes.add_textbox(
            left_margin + Emu(100000), y_offset + Emu(300000),
            company_label_width, Emu(200000)
        )
        ctf = count_box.text_frame
        cp = ctf.paragraphs[0]
        cp.text = f"{len(arts)}건"
        cp.font.size = Pt(10)
        cp.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        cp.font.name = "맑은 고딕"

        # 기사 내용 (bullet)
        content_box = slide.shapes.add_textbox(
            left_margin + company_label_width + Emu(100000), y_offset,
            content_width, Emu(900000)
        )
        ctf = content_box.text_frame
        ctf.word_wrap = True

        if not arts:
            p = ctf.paragraphs[0]
            p.text = "해당 기간 주요 동향 없음"
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
            p.font.name = "맑은 고딕"
            p.font.italic = True
        else:
            for i, art in enumerate(arts[:5]):
                if i == 0:
                    p = ctf.paragraphs[0]
                else:
                    p = ctf.add_paragraph()
                title = art.get("title", "")
                if len(title) > 70:
                    title = title[:67] + "..."
                p.text = f"▸ {title}"
                p.font.size = Pt(10)
                p.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
                p.font.name = "맑은 고딕"
                p.space_after = Pt(3)
                p.line_spacing = Pt(16)

                # 날짜 추가
                run = p.add_run()
                run.text = f"  ({art.get('date', '')})"
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
                run.font.name = "맑은 고딕"

        # 구분선
        sep_y = y_offset + Emu(920000)
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left_margin, sep_y, block_width, Emu(12000)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0xE8, 0xE8, 0xE8)
        line.line.fill.background()

        y_offset = sep_y + Emu(80000)

    # ── 하단 푸터 ──
    footer_box = slide.shapes.add_textbox(
        Emu(400000), prs.slide_height - Emu(400000),
        Emu(11000000), Emu(250000)
    )
    ftf = footer_box.text_frame
    fp = ftf.paragraphs[0]
    fp.text = f"대형사 동향 · 네이버 뉴스 API 기반 사업동향 자동 수집 · {now.strftime('%Y.%m.%d')}"
    fp.font.size = Pt(8)
    fp.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)
    fp.font.name = "맑은 고딕"
    fp.alignment = PP_ALIGN.CENTER

    prs.save(PPT_PATH)
    print(f"  → PPT 저장 완료: {PPT_PATH}")


def main():
    print(f"\n{'='*50}")
    print(f"  대형사 동향 PPT 보고서 생성")
    print(f"{'='*50}\n")

    print("[1/2] CSV 데이터 로드 중...")
    company_articles = load_articles()
    for name, arts in company_articles.items():
        print(f"  [{name}] {len(arts)}건")

    print("[2/2] PPT 생성 중...")
    create_report(company_articles)

    print(f"\n  완료!")


if __name__ == "__main__":
    main()
