import os
import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor

def update_presentation():
    src_path = r"C:\Users\haiha\.gemini\antigravity\brain\629856cd-4ccf-4073-ba60-004fe2943b05\.user_uploaded\media_1790768850898.pptx"
    dst_path = r"C:\Coding\Jungle\5Week\5Week_Troubleshooting.pptx"

    prs = pptx.Presentation(src_path)

    # Colors for syntax highlighting in code boxes (code box is dark)
    CLR_KEYWORD = RGBColor(255, 123, 114)   # Salmon red (int)
    CLR_FUNC = RGBColor(207, 165, 251)      # Lilac purple (malloc, memset, memcpy)
    CLR_PAREN = RGBColor(255, 215, 0)       # Gold ((, ))
    CLR_NUM = RGBColor(180, 204, 167)       # Sage green (0, 5, 10, ...)
    CLR_SIZEOF = RGBColor(78, 155, 213)     # Blue (sizeof)
    CLR_COMMENT = RGBColor(139, 148, 158)   # Muted gray (/* ... */, // ...)
    CLR_CODE_TXT = RGBColor(240, 240, 240)  # White text inside dark code box

    # Explanation bullets color (outer card is light, so text MUST be black!)
    CLR_EXPLAIN = RGBColor(0, 0, 0)         # Pure Black

    def set_rich_bullets(shape, paragraph_runs, font_size_pt=12.0):
        tf = shape.text_frame
        tf.word_wrap = True
        tf.clear()
        
        for p_idx, runs in enumerate(paragraph_runs):
            p = tf.add_paragraph() if p_idx > 0 else tf.paragraphs[0]
            if not runs:  # Empty line spacer
                p.text = ""
                continue
            for text, is_bold in runs:
                r = p.add_run()
                r.text = text
                r.font.name = "Malgun Gothic"
                r.font.size = Pt(font_size_pt)
                r.font.bold = is_bold
                r.font.color.rgb = CLR_EXPLAIN

    def adjust_textbox_pos(shape, left=None, top=None, width=None, height=None):
        if left is not None:
            shape.left = left
        if top is not None:
            shape.top = top
        if width is not None:
            shape.width = width
        if height is not None:
            shape.height = height

    # Remove empty leftover shape 30 (TextBox 29) on any slide
    for slide in prs.slides:
        for sh in list(slide.shapes):
            if sh.shape_id == 30:
                slide.shapes._spTree.remove(sh._element)

    # ==========================================
    # SLIDE 1: malloc & calloc
    # ==========================================
    s1 = prs.slides[0]
    
    # Header & Titles bold verification
    tb_s1_header = [sh for sh in s1.shapes if sh.shape_id == 5][0]
    for r in tb_s1_header.text_frame.paragraphs[0].runs:
        r.font.bold = True
    tb_s1_t_l = [sh for sh in s1.shapes if sh.shape_id == 56][0]
    for r in tb_s1_t_l.text_frame.paragraphs[0].runs:
        r.font.bold = True
    tb_s1_t_r = [sh for sh in s1.shapes if sh.shape_id == 15][0]
    for r in tb_s1_t_r.text_frame.paragraphs[0].runs:
        r.font.bold = True

    tb_malloc = [sh for sh in s1.shapes if sh.shape_id == 10][0]
    tb_calloc = [sh for sh in s1.shapes if sh.shape_id == 24][0]

    adjust_textbox_pos(tb_malloc, left=1114975, top=4450000, width=4241666, height=1850000)
    adjust_textbox_pos(tb_calloc, left=6835360, top=4450000, width=4241666, height=1850000)

    malloc_runs = [
        [("• 원하는 크기만큼 ", False), ("힙(Heap) 메모리", True), ("를 빌려오는 가장 기본적이고 빠른 함수임.", False)],
        [("    + ", False), ("빌려오기만 하면 끝이야? 컴퓨터가 알아서 다 준비해 줘?", False)],
        [],
        [("• 공간만 빌려줄 뿐 청소를 안 해줘서, ", False), ("이전에 쓰던 쓰레기 값", True), ("이 그대로 남아있음.", False)],
        [("• 값을 초기화하지 않고 읽거나 포인터로 쓰면 ", False), ("SIGSEGV 크래시 발생!", True)]
    ]
    set_rich_bullets(tb_malloc, malloc_runs, 12.0)

    calloc_runs = [
        [("• 메모리를 할당받자마자 ", False), ("모든 비트를 0(NULL)으로 깨끗하게 초기화", True), ("해 줌.", False)],
        [("    + ", False), ("malloc 쓰고 나서 내가 직접 0 넣는 거랑 뭐가 달라?", False)],
        [],
        [("• 할당과 0 청소가 한 번에 이루어져, ", False), ("초기화 누락 버그를 원천 차단", True), ("함.", False)],
        [("    + ", False), ("그럼 malloc 대신 항상 calloc만 쓰면 되는 거 아냐?", False)],
        [],
        [("• 모든 메모리를 0으로 미는 작업 때문에 ", False), ("malloc보다 미세하게 느림", True), (". 안전이 우선일 때 사용!", False)]
    ]
    set_rich_bullets(tb_calloc, calloc_runs, 12.0)

    # ==========================================
    # SLIDE 2: realloc & free
    # ==========================================
    s2 = prs.slides[1]

    # Header & Titles bold verification
    tb_s2_header = [sh for sh in s2.shapes if sh.shape_id == 5][0]
    for r in tb_s2_header.text_frame.paragraphs[0].runs:
        r.font.bold = True
    tb_s2_t_l = [sh for sh in s2.shapes if sh.shape_id == 56][0]
    tb_s2_t_l.text_frame.paragraphs[0].text = "realloc() – Re-Allocation"
    for r in tb_s2_t_l.text_frame.paragraphs[0].runs:
        r.font.bold = True
    tb_s2_t_r = [sh for sh in s2.shapes if sh.shape_id == 15][0]
    tb_s2_t_r.text_frame.paragraphs[0].text = "free() – Memory Free"
    for r in tb_s2_t_r.text_frame.paragraphs[0].runs:
        r.font.bold = True

    # Fix realloc syntax: change ", " to " * "
    code_realloc = [sh for sh in s2.shapes if sh.shape_id == 57][0]
    p0 = code_realloc.text_frame.paragraphs[0]
    for r in p0.runs:
        if r.text == ", ":
            r.text = " * "

    tb_realloc = [sh for sh in s2.shapes if sh.shape_id == 10][0]
    tb_free = [sh for sh in s2.shapes if sh.shape_id == 24][0]

    adjust_textbox_pos(tb_realloc, left=1114975, top=4480000, width=4241666, height=1850000)
    adjust_textbox_pos(tb_free, left=6835360, top=4480000, width=4241666, height=1850000)

    realloc_runs = [
        [("• 이미 빌려둔 ", False), ("메모리의 크기를 늘리거나 줄일 때", True), (" 사용함.", False)],
        [("    + ", False), ("그냥 살던 방 벽 터서 제자리에서 늘리면 되는 거 아냐?", False)],
        [],
        [("• 옆 방에 다른 데이터가 살고 있으면 ", False), ("짐을 몽땅 새 집으로 옮겨주고 예전 집은 폭파(free)", True), ("함.", False)],
        [("    + ", False), ("그럼 왜 ", False), ("arr = new_arr", True), ("로 대입해? new_arr에 arr을 넣어야 하는 거 아냐?", False)],
        [],
        [("• realloc이 이미 새 집(new_arr)에 짐을 다 옮겨뒀기 때문에, ", False), ("주소록(arr)만 새 주소로 갱신", True), ("하면 됨!", False)]
    ]
    set_rich_bullets(tb_realloc, realloc_runs, 12.0)

    free_runs = [
        [("• 다 쓴 동적 메모리를 ", False), ("운영체제(OS)에 반납", True), ("함.", False)],
        [("    + ", False), ("반납 안 하면 어떻게 되는데? 컴퓨터 끄면 어차피 사라지잖아?", False)],
        [],
        [("• 프로그램 실행 중 반납하지 않은 메모리는 계속 누적되어 ", False), ("메모리 누수(Memory Leak)", True), ("가 발생함.", False)],
        [("    + ", False), ("그럼 free하고 나서 왜 굳이 ", False), ("arr = NULL", True), ("을 해?", False)],
        [],
        [("• 방을 뺐는데 주소록을 안 지우면, 이미 반납한 방을 또 반납(", False), ("Double Free", True), (")하거나 잘못 찾아가(", False), ("Use-After-Free", True), (") 강제 종료됨.", False)]
    ]
    set_rich_bullets(tb_free, free_runs, 12.0)

    # ==========================================
    # SLIDE 3: memset & memcpy
    # ==========================================
    s3 = prs.slides[2]

    # Header bold & styling matching Slide 1 & 2
    tb_s3_header = [sh for sh in s3.shapes if sh.shape_id == 5][0]
    p_hdr = tb_s3_header.text_frame.paragraphs[0]
    p_hdr.text = "트러블 슈팅 – 메모리 조작 함수"
    p_hdr.font.name = "+mj-lt"
    p_hdr.font.size = Pt(24)
    p_hdr.font.bold = True
    for r in p_hdr.runs:
        r.font.bold = True

    # Slide 3 Titles: bold=True
    tb_s3_t_l = [sh for sh in s3.shapes if sh.shape_id == 56][0]
    p_l = tb_s3_t_l.text_frame.paragraphs[0]
    p_l.text = "memset() – Memory Set"
    p_l.font.bold = True
    for r in p_l.runs:
        r.font.bold = True

    tb_s3_t_r = [sh for sh in s3.shapes if sh.shape_id == 15][0]
    p_r = tb_s3_t_r.text_frame.paragraphs[0]
    p_r.text = "memcpy() – Memory Copy"
    p_r.font.bold = True
    for r in p_r.runs:
        r.font.bold = True

    # Slide 3 Left: memset code
    code_memset = [sh for sh in s3.shapes if sh.shape_id == 57][0]
    # Align position to match Slide 1 code box!
    adjust_textbox_pos(code_memset, left=1114975, top=2740502, width=4320540, height=1324768)

    tf_memset = code_memset.text_frame
    tf_memset.clear()

    # Paragraph 0: int arr[10];
    p0 = tf_memset.paragraphs[0]
    r = p0.add_run(); r.text = "int"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_KEYWORD
    r = p0.add_run(); r.text = " arr["; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT
    r = p0.add_run(); r.text = "10"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_NUM
    r = p0.add_run(); r.text = "];"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT

    # Paragraph 1: memset (arr, 0, sizeof(arr));
    p1 = tf_memset.add_paragraph()
    r = p1.add_run(); r.text = "memset "; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_FUNC
    r = p1.add_run(); r.text = "("; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_PAREN
    r = p1.add_run(); r.text = "arr, "; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT
    r = p1.add_run(); r.text = "0"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_NUM
    r = p1.add_run(); r.text = ", "; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT
    r = p1.add_run(); r.text = "sizeof"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_SIZEOF
    r = p1.add_run(); r.text = "("; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_PAREN
    r = p1.add_run(); r.text = "arr"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT
    r = p1.add_run(); r.text = "));"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_PAREN

    # Paragraph 2: empty
    p2 = tf_memset.add_paragraph()

    # Paragraph 3: /*
    p3 = tf_memset.add_paragraph()
    r = p3.add_run(); r.text = "/* "; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_COMMENT

    # Paragraph 4: comment
    p4 = tf_memset.add_paragraph()
    r = p4.add_run(); r.text = "arr의 메모리 공간을 0으로 초기화 (40byte)"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_COMMENT

    # Paragraph 5: comment
    p5 = tf_memset.add_paragraph()
    r = p5.add_run(); r.text = "원하는 크기만큼 특정 바이트 값으로 채움"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_COMMENT

    # Paragraph 6: */
    p6 = tf_memset.add_paragraph()
    r = p6.add_run(); r.text = "*/"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_COMMENT

    # Slide 3 Right: memcpy code
    code_memcpy = [sh for sh in s3.shapes if sh.shape_id == 58][0]
    # Align position to match Slide 1 code box!
    adjust_textbox_pos(code_memcpy, left=6786569, top=2740502, width=4320540, height=1324768)

    tf_memcpy = code_memcpy.text_frame
    tf_memcpy.clear()

    # Paragraph 0: int src[5] = {1, 2, 3, 4, 5};
    p0 = tf_memcpy.paragraphs[0]
    r = p0.add_run(); r.text = "int"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_KEYWORD
    r = p0.add_run(); r.text = " src["; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT
    r = p0.add_run(); r.text = "5"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_NUM
    r = p0.add_run(); r.text = "] = {"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT
    r = p0.add_run(); r.text = "1, 2, 3, 4, 5"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_NUM
    r = p0.add_run(); r.text = "};"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT

    # Paragraph 1: int dest[5];
    p1 = tf_memcpy.add_paragraph()
    r = p1.add_run(); r.text = "int"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_KEYWORD
    r = p1.add_run(); r.text = " dest["; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT
    r = p1.add_run(); r.text = "5"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_NUM
    r = p1.add_run(); r.text = "];"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT

    # Paragraph 2: memcpy (dest, src, sizeof(src));
    p2 = tf_memcpy.add_paragraph()
    r = p2.add_run(); r.text = "memcpy "; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_FUNC
    r = p2.add_run(); r.text = "("; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_PAREN
    r = p2.add_run(); r.text = "dest, src, "; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT
    r = p2.add_run(); r.text = "sizeof"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_SIZEOF
    r = p2.add_run(); r.text = "("; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_PAREN
    r = p2.add_run(); r.text = "src"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_CODE_TXT
    r = p2.add_run(); r.text = "));"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_PAREN

    # Paragraph 3: /*
    p3 = tf_memcpy.add_paragraph()
    r = p3.add_run(); r.text = "/* "; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_COMMENT

    # Paragraph 4: comment
    p4 = tf_memcpy.add_paragraph()
    r = p4.add_run(); r.text = "src의 내용을 dest로 그대로 복사 (Deep Copy)"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_COMMENT

    # Paragraph 5: comment
    p5 = tf_memcpy.add_paragraph()
    r = p5.add_run(); r.text = "지정한 크기(20byte)만큼 고속 메모리 복사"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_COMMENT

    # Paragraph 6: */
    p6 = tf_memcpy.add_paragraph()
    r = p6.add_run(); r.text = "*/"; r.font.name = "Consolas"; r.font.size = Pt(14); r.font.color.rgb = CLR_COMMENT

    # Bullets on Slide 3
    tb_memset = [sh for sh in s3.shapes if sh.shape_id == 10][0]
    tb_memcpy = [sh for sh in s3.shapes if sh.shape_id == 24][0]

    adjust_textbox_pos(tb_memset, left=1114975, top=4450000, width=4241666, height=1850000)
    adjust_textbox_pos(tb_memcpy, left=6835360, top=4450000, width=4241666, height=1850000)

    memset_runs = [
        [("• 특정 메모리 구간을 원하는 값(주로 0)으로 ", False), ("한 번에 도배(초기화)", True), ("함.", False)],
        [("    + ", False), ("calloc도 0으로 채워주는데 memset이 왜 따로 필요해?", False)],
        [],
        [("• calloc은 힙(동적 메모리) 전용이지만, memset은 ", False), ("스택(지역 변수), 전역 변수 어디든", True), (" 0으로 밀어버릴 수 있음.", False)],
        [("    + ", False), ("그럼 memset으로 int 배열을 1로 다 채울 수도 있어?", False)],
        [],
        [("• ", False), ("1바이트 단위", True), ("로 칠하기 때문에 int에 1을 넣으면 ", False), ("0x01010101(엉뚱한 큰 수)", True), ("가 됨! 0으로 밀 때만 안전함.", False)]
    ]
    set_rich_bullets(tb_memset, memset_runs, 12.0)

    memcpy_runs = [
        [("• 특정 메모리 영역의 데이터를 다른 곳으로 ", False), ("고속 복사(Deep Copy)", True), ("함.", False)],
        [("    + ", False), ("그냥 ", False), ("dest = src", True), (" 이렇게 포인터 대입하면 안 돼?", False)],
        [],
        [("• 'dest = src'는 주소만 넘기는 얕은 복사지만, memcpy는 ", False), ("완전히 독립된 새 데이터 공간", True), ("을 만듦.", False)],
        [("    + ", False), ("복사할 크기를 잘못 적으면 어떻게 돼?", False)],
        [],
        [("• 원본보다 크게 적으면 남의 방까지 침범하는 ", False), ("버퍼 오버플로우(Buffer Overflow)", True), ("가 나고, 덜 적으면 데이터가 잘려 나감.", False)]
    ]
    set_rich_bullets(tb_memcpy, memcpy_runs, 12.0)

    prs.save(dst_path)
    print("Successfully updated presentation with black text, bold highlights, and adjusted positions!")

if __name__ == "__main__":
    update_presentation()
