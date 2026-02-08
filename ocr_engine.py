import sys
import json
import os
import cv2
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')


def process_image(image_path):
    """순수 OpenCV 방식으로 테이블 색상 분석"""

    print(f"이미지 분석 시작: {image_path}", file=sys.stderr)
    print("방식: 순수 OpenCV (그리드 라인 감지)", file=sys.stderr)

    # 이미지 로드
    img = cv2.imread(image_path)
    if img is None:
        raise Exception("이미지를 로드할 수 없습니다.")

    height, width = img.shape[:2]
    print(f"이미지 크기: {width} x {height}", file=sys.stderr)

    # 그레이스케일 변환
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 엣지 감지
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)

    # 수평/수직 라인 감지
    horizontal_lines = detect_lines(edges, 'horizontal', width, height)
    vertical_lines = detect_lines(edges, 'vertical', width, height)

    print(f"감지된 라인: 수평 {len(horizontal_lines)}개, 수직 {len(vertical_lines)}개", file=sys.stderr)

    # 라인이 부족하면 기본 그리드 사용
    if len(horizontal_lines) < 3 or len(vertical_lines) < 3:
        print("라인 감지 부족, 기본 그리드 사용", file=sys.stderr)
        results = analyze_with_default_grid(img, width, height)
    else:
        # 그리드 셀 분석
        results = analyze_grid_cells(img, horizontal_lines, vertical_lines, width, height)

    # 결과 정보
    if len(results) > 0:
        first_floor = results[0]
        num_units = len(first_floor.get('units', {}))
        print(f"감지된 구조: {len(results)}층 × {num_units}호", file=sys.stderr)

    # 색상 분포
    color_count = {'GREEN': 0, 'YELLOW': 0, 'PINK': 0, 'WHITE': 0}
    for floor_data in results:
        for unit_key, unit_data in floor_data['units'].items():
            color = unit_data['color']
            color_count[color] = color_count.get(color, 0) + 1

    print(f"색상 분포: {color_count}", file=sys.stderr)

    return results


def detect_lines(edges, direction, width, height):
    """수평 또는 수직 라인 감지"""

    lines = cv2.HoughLinesP(
        edges,
        rho=1,
        theta=np.pi/180,
        threshold=100,
        minLineLength=min(width, height) * 0.3,
        maxLineGap=10
    )

    if lines is None:
        return []

    result = []
    for line in lines:
        x1, y1, x2, y2 = line[0]

        if direction == 'horizontal':
            # 수평 라인: y값 차이가 작음
            if abs(y2 - y1) < 10:
                y_avg = (y1 + y2) // 2
                result.append(y_avg)
        else:
            # 수직 라인: x값 차이가 작음
            if abs(x2 - x1) < 10:
                x_avg = (x1 + x2) // 2
                result.append(x_avg)

    # 중복 제거 및 정렬
    result = sorted(set(result))

    # 너무 가까운 라인 병합 (10픽셀 이내)
    merged = []
    for val in result:
        if not merged or val - merged[-1] > 15:
            merged.append(val)

    return merged


def analyze_grid_cells(img, h_lines, v_lines, width, height):
    """감지된 그리드를 기반으로 셀 분석"""

    # 헤더 제외 (첫 번째 수평 라인 이후부터)
    if len(h_lines) > 1:
        data_h_lines = h_lines[1:]  # 헤더 행 제외
    else:
        data_h_lines = h_lines

    # 층 열 제외 (첫 번째 수직 라인 이후부터)
    if len(v_lines) > 1:
        data_v_lines = v_lines[1:]  # 층 열 제외
    else:
        data_v_lines = v_lines

    num_rows = len(data_h_lines) - 1 if len(data_h_lines) > 1 else 25
    num_cols = len(data_v_lines) - 1 if len(data_v_lines) > 1 else 10

    print(f"그리드: {num_rows}행 x {num_cols}열", file=sys.stderr)

    results = []

    for row in range(num_rows):
        floor_num = num_rows - row
        floor_data = {
            "floor": f"{floor_num}층",
            "units": {}
        }

        for col in range(num_cols):
            unit_num = col + 1

            # 셀 경계 계산
            if row < len(data_h_lines) - 1 and col < len(data_v_lines) - 1:
                y1 = data_h_lines[row]
                y2 = data_h_lines[row + 1]
                x1 = data_v_lines[col]
                x2 = data_v_lines[col + 1]
            else:
                continue

            # 셀 중심 영역에서 색상 샘플링 (테두리 제외)
            margin_x = int((x2 - x1) * 0.2)
            margin_y = int((y2 - y1) * 0.2)

            sx1 = x1 + margin_x
            sy1 = y1 + margin_y
            sx2 = x2 - margin_x
            sy2 = y2 - margin_y

            # 경계 확인
            sx1 = max(0, min(sx1, width - 1))
            sy1 = max(0, min(sy1, height - 1))
            sx2 = max(sx1 + 1, min(sx2, width))
            sy2 = max(sy1 + 1, min(sy2, height))

            # 색상 측정
            roi = img[sy1:sy2, sx1:sx2]
            if roi.size > 0:
                avg_color = cv2.mean(roi)[:3]
                b, g, r = avg_color
                color = classify_color(r, g, b)
            else:
                color = "WHITE"

            floor_data["units"][f"{unit_num}호"] = {
                "text": "",
                "color": color
            }

        if floor_data["units"]:
            results.append(floor_data)

    return results


def analyze_with_default_grid(img, width, height):
    """기본 그리드로 분석 (라인 감지 실패 시)"""

    # 테이블 영역 추정 (이미지의 중앙 부분)
    # 상단 15%, 하단 5%, 좌우 5% 여백
    margin_top = int(height * 0.15)
    margin_bottom = int(height * 0.05)
    margin_left = int(width * 0.08)
    margin_right = int(width * 0.02)

    table_x1 = margin_left
    table_y1 = margin_top
    table_x2 = width - margin_right
    table_y2 = height - margin_bottom

    table_width = table_x2 - table_x1
    table_height = table_y2 - table_y1

    # 기본 그리드: 25행 x 10열
    rows = 25
    cols = 10

    cell_width = table_width / (cols + 1)  # +1 for 층 column
    cell_height = table_height / (rows + 1)  # +1 for header

    print(f"기본 그리드 사용: {rows}행 x {cols}열", file=sys.stderr)
    print(f"테이블 영역: ({table_x1}, {table_y1}) ~ ({table_x2}, {table_y2})", file=sys.stderr)

    results = []

    for row in range(rows):
        floor_num = rows - row
        floor_data = {
            "floor": f"{floor_num}층",
            "units": {}
        }

        for col in range(cols):
            unit_num = col + 1

            # 셀 중심 좌표 (헤더와 층 열 건너뛰기)
            cx = int(table_x1 + (col + 1.5) * cell_width)
            cy = int(table_y1 + (row + 1.5) * cell_height)

            # 샘플링 영역
            sample_size = int(min(cell_width, cell_height) * 0.25)
            sx1 = max(0, cx - sample_size)
            sy1 = max(0, cy - sample_size)
            sx2 = min(width, cx + sample_size)
            sy2 = min(height, cy + sample_size)

            # 색상 측정
            roi = img[sy1:sy2, sx1:sx2]
            if roi.size > 0:
                avg_color = cv2.mean(roi)[:3]
                b, g, r = avg_color
                color = classify_color(r, g, b)
            else:
                color = "WHITE"

            floor_data["units"][f"{unit_num}호"] = {
                "text": "",
                "color": color
            }

        results.append(floor_data)

    return results


def classify_color(r, g, b):
    """RGB 값을 기반으로 색상 분류 - 아파트 현황표 전용"""

    # 디버그 출력 (필요시 활성화)
    # print(f"RGB({r:.0f},{g:.0f},{b:.0f})", file=sys.stderr)

    # 1. 직접 RGB 비교 방식 (파스텔 색상에 효과적)

    # 흰색 판정: 모든 채널이 높고 비슷함
    if r > 245 and g > 245 and b > 245:
        return "WHITE"

    # 밝기
    brightness = (r + g + b) / 3

    # 채널 차이
    rg_diff = abs(r - g)
    rb_diff = abs(r - b)
    gb_diff = abs(g - b)

    # 거의 흰색 (밝고 채널 차이 작음)
    if brightness > 240 and max(rg_diff, rb_diff, gb_diff) < 15:
        return "WHITE"

    # YELLOW 판정: R과 G가 높고, B가 낮음
    # 연한 노란색: R≈G > B
    if r > 220 and g > 220 and b < 230:
        if r > b + 10 and g > b + 10:
            if abs(r - g) < 30:  # R과 G가 비슷
                return "YELLOW"

    # GREEN 판정: G가 가장 높음
    # 연한 녹색: G > R, G > B
    if g > 200:
        if g > r + 5 and g > b + 5:
            return "GREEN"

    # PINK 판정: R이 높고, B도 상대적으로 높음
    # 연한 분홍: R > G, B > G 또는 R≈B > G
    if r > 220 and b > 200:
        if r > g and b > g - 20:
            if r >= b - 30:  # R과 B가 비슷하거나 R이 더 높음
                return "PINK"

    # 2. HSV 기반 추가 판정
    r_n, g_n, b_n = r / 255.0, g / 255.0, b / 255.0
    max_c = max(r_n, g_n, b_n)
    min_c = min(r_n, g_n, b_n)
    diff = max_c - min_c

    v = max_c
    s = 0 if max_c == 0 else diff / max_c

    if diff == 0:
        h = 0
    elif max_c == r_n:
        h = 60 * (((g_n - b_n) / diff) % 6)
    elif max_c == g_n:
        h = 60 * (((b_n - r_n) / diff) + 2)
    else:
        h = 60 * (((r_n - g_n) / diff) + 4)

    # 채도가 매우 낮으면 흰색
    if s < 0.05:
        return "WHITE"

    # HSV 기반 색상 분류
    if s > 0.05:
        # 노란색 영역: 40-70도
        if 40 <= h <= 70:
            return "YELLOW"

        # 녹색 영역: 70-160도
        if 70 < h <= 160:
            return "GREEN"

        # 분홍/빨강 영역: 300-360 또는 0-30도
        if h > 300 or h < 35:
            return "PINK"

        # 보라/분홍 영역: 260-300도
        if 260 <= h <= 330:
            return "PINK"

    return "WHITE"


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "이미지 경로가 필요합니다."}))
        sys.exit(1)

    image_path = sys.argv[1]

    try:
        data = process_image(image_path)
        print(json.dumps(data, ensure_ascii=False))
    except Exception as e:
        import traceback
        traceback.print_exc(file=sys.stderr)
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
