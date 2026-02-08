import sys
import json
import cv2
import numpy as np
import os

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# EasyOCR 초기화 (텍스트 인식용)
try:
    import easyocr
    OCR_READER = easyocr.Reader(['en'], gpu=False, verbose=False)
    OCR_AVAILABLE = True
    print("EasyOCR 초기화 완료", file=sys.stderr)
except ImportError:
    OCR_READER = None
    OCR_AVAILABLE = False
    print("EasyOCR 없음 - 텍스트 인식 비활성화", file=sys.stderr)

# ============================================================
# 설정 (고정 그리드 - LG신주례 현황표 기준)
# ============================================================
NUM_FLOORS = 25  # 층수 (25층~1층)
NUM_UNITS = 10   # 호수 (1호~10호)


def load_image(image_path):
    """이미지 로드 (한글 경로 지원)"""
    img_array = np.fromfile(image_path, np.uint8)
    img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    if img is None:
        raise Exception(f"이미지 로드 실패: {image_path}")
    return img


def detect_table_grid(img, colored_regions=None):
    """표의 그리드 라인 검출 (정밀화 버전)"""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # 노이즈 제거 및 대비 향상
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    # 어댑티브 threshold
    binary = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                 cv2.THRESH_BINARY_INV, 15, 3)
    
    # 커널 사이즈를 조금 더 작게 조정하여 얇은 선들도 검출
    img_h, img_w = img.shape[:2]
    h_size = img_w // 30
    v_size = img_h // 40
    
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_size, 1))
    horizontal_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
    
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_size))
    vertical_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel)
    
    # HoughLinesP 파라미터 완화
    h_lines = cv2.HoughLinesP(horizontal_lines, 1, np.pi/180, threshold=50, minLineLength=h_size, maxLineGap=20)
    v_lines = cv2.HoughLinesP(vertical_lines, 1, np.pi/180, threshold=50, minLineLength=v_size, maxLineGap=20)
    
    def group_coords(coords, gap=5):
        if not coords: return []
        coords = sorted(coords)
        groups = []
        if coords:
            curr_group = [coords[0]]
            for i in range(1, len(coords)):
                if coords[i] - coords[i-1] <= gap:
                    curr_group.append(coords[i])
                else:
                    groups.append(int(np.mean(curr_group)))
                    curr_group = [coords[i]]
            groups.append(int(np.mean(curr_group)))
        return groups

    h_coords_raw = []
    if h_lines is not None:
        for line in h_lines:
            x1, y1, x2, y2 = line[0]
            # 수평에 가깝고 길이가 어느 정도 있는 선만 수집
            line_len = abs(x1 - x2)
            if abs(y1 - y2) < 5 and line_len > (img_w // 4):
                h_coords_raw.append((y1 + y2) // 2)
    
    v_coords_raw = []
    if v_lines is not None:
        for line in v_lines:
            x1, y1, x2, y2 = line[0]
            # 수직에 가깝고 길이가 어느 정도 있는 선만 수집
            line_len = abs(y1 - y2)
            if abs(x1 - x2) < 5 and line_len > (img_h // 10):
                v_coords_raw.append((x1 + x2) // 2)
    
    # 근접한 선들 그룹화 (더 타이트하게)
    h_coords = group_coords(h_coords_raw, gap=5)
    v_coords = group_coords(v_coords_raw, gap=5)
    
    # ------------------------------------------------------------
    # 그리드 보간 (Missing Lines Fill) - 색상 영역 간 거리 기반
    # ------------------------------------------------------------
    def interpolate_lines_smart(coords, regions, axis='y'):
        if len(coords) < 2: return coords
        
        # 실제 데이터 점들 사이의 거리로부터 셀 크기 추정 (정교화)
        actual_gaps = []
        if regions:
            if axis == 'x': # 열 너비 추정 (같은 행에 있는 점들끼리 비교)
                # y좌표 기준으로 정렬
                check_regions = sorted(regions, key=lambda r: r['center'][1])
                for i in range(len(check_regions)):
                    for j in range(i+1, min(i+10, len(check_regions))):
                        r1, r2 = check_regions[i], check_regions[j]
                        # Y좌표 차이가 작으면 같은 행으로 간주
                        if abs(r1['center'][1] - r2['center'][1]) < 15:
                            gap = abs(r1['center'][0] - r2['center'][0])
                            if 40 < gap < 200: actual_gaps.append(gap) # 최소 40px 이상
            
            else: # 행 높이 추정 (같은 열에 있는 점들끼리 비교)
                # x좌표 기준으로 정렬
                check_regions = sorted(regions, key=lambda r: r['center'][0])
                for i in range(len(check_regions)):
                    for j in range(i+1, min(i+10, len(check_regions))):
                        r1, r2 = check_regions[i], check_regions[j]
                        # X좌표 차이가 작으면 같은 열로 간주
                        if abs(r1['center'][0] - r2['center'][0]) < 15:
                            gap = abs(r1['center'][1] - r2['center'][1])
                            if 15 < gap < 100: actual_gaps.append(gap)
        
        if actual_gaps:
            # 데이터 간격과 라인 간격의 중앙값 비교
            median_actual = np.median(actual_gaps)
            
            line_gaps = [coords[i+1] - coords[i] for i in range(len(coords)-1)]
            median_line = np.median(line_gaps) if line_gaps else 0
            
            if median_line > 0:
                # 1. 라인 간격이 너무 작으면 (잡음) -> 데이터 간격 신뢰 (예: X축)
                is_line_noise = median_line < median_actual * 0.6
                # 2. 데이터 간격이 너무 크면 (희소) -> 라인 간격 신뢰 (예: Y축)
                is_data_sparse = median_actual > median_line * 1.8
                
                # 우선순위 조정: Y축(행)은 데이터가 희소하여 actual이 뻥튀기될 수 있으므로 sparse 체크 우선
                if axis == 'y' and is_data_sparse:
                    estimated_cell_size = median_line
                elif is_line_noise:
                    estimated_cell_size = median_actual
                elif is_data_sparse:
                    estimated_cell_size = median_line
                else:
                    estimated_cell_size = median_actual
            else:
                estimated_cell_size = median_actual
                
            # 최소 안전장치
            estimated_cell_size = max(estimated_cell_size, 10)
        else:
            line_gaps = [coords[i+1] - coords[i] for i in range(len(coords)-1)]
            estimated_cell_size = np.median(line_gaps) if line_gaps else 20

        print(f"추정된 {axis}축 셀 크기: {estimated_cell_size:.1f} (데이터:{median_actual:.1f}, 라인:{median_line:.1f})", file=sys.stderr)
        
        new_coords = [coords[0]]
        for i in range(len(coords)-1):
            gap = coords[i+1] - coords[i]
            # 너무 촘촘한 선 제거 (셀 크기의 0.5배 미만 - 기준 완화)
            if gap < estimated_cell_size * 0.5:
                continue 
            
            # 누락된 선 보간 (셀 크기의 1.5배 이상)
            if gap > estimated_cell_size * 1.5:
                num_missing = int(round(gap / estimated_cell_size)) - 1
                for j in range(1, num_missing + 1):
                    new_coords.append(int(coords[i] + j * (gap / (num_missing + 1))))
            
            new_coords.append(coords[i+1])
            
        return sorted(list(set(new_coords))), estimated_cell_size

    h_coords, h_size = interpolate_lines_smart(h_coords, colored_regions, 'y')
    v_coords, v_size = interpolate_lines_smart(v_coords, colored_regions, 'x')
    
    # 보간 후 중복/근접 라인 재병합 (중요) - 셀 크기의 50% 이내 라인은 하나로 합침
    if h_coords and len(h_coords) > 1:
        h_coords = group_coords(h_coords, gap=h_size * 0.5)
        
    if v_coords and len(v_coords) > 1:
        v_coords = group_coords(v_coords, gap=v_size * 0.5)
    
    # 너무 좁은 간격 다시 필터링 (최소 셀 크기 강제 적용)
    def filter_close_lines(coords, min_gap):
        if not coords: return []
        filtered = [coords[0]]
        for i in range(1, len(coords)):
            if coords[i] - filtered[-1] >= min_gap:
                filtered.append(coords[i])
        return filtered

    h_coords = filter_close_lines(h_coords, h_size * 0.8)
    v_coords = filter_close_lines(v_coords, v_size * 0.8)
    
    print(f"최종 수평선: {len(h_coords)}개, 수직선: {len(v_coords)}개", file=sys.stderr)
    
    return h_coords, v_coords


def find_all_colored_regions(img):
    """전체 이미지에서 색상 영역 검출 (파스텔 톤 최적화)"""
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)

    # 색상별 HSV 범위 (파스텔 색상에 맞게 조정)
    color_ranges = {
        'GREEN': [
            ([35, 30, 150], [85, 255, 255]),     # 연한 녹색 (파스텔)
            ([35, 50, 100], [85, 255, 255]),     # 중간 녹색
            ([40, 20, 180], [80, 150, 255]),     # 매우 연한 녹색
        ],
        'YELLOW': [
            ([15, 30, 200], [35, 255, 255]),     # 연한 노랑 (파스텔)
            ([20, 50, 180], [35, 200, 255]),     # 중간 노랑
            ([18, 20, 220], [32, 120, 255]),     # 매우 연한 노랑
        ],
        'PINK': [
            ([140, 20, 180], [180, 150, 255]),   # 연한 분홍 (파스텔)
            ([150, 30, 200], [170, 120, 255]),   # 중간 분홍
            ([0, 20, 200], [15, 100, 255]),      # 빨강 계열 연한 분홍
            ([160, 15, 200], [180, 80, 255]),    # 매우 연한 분홍
        ],
    }

    all_regions = []

    for color_name, ranges in color_ranges.items():
        combined_mask = np.zeros(img.shape[:2], dtype=np.uint8)

        for lower, upper in ranges:
            lower = np.array(lower)
            upper = np.array(upper)
            mask = cv2.inRange(hsv, lower, upper)
            combined_mask = cv2.bitwise_or(combined_mask, mask)

        # 노이즈 제거
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
        combined_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_OPEN, kernel)
        combined_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_CLOSE, kernel)

        # 윤곽선 찾기
        contours, _ = cv2.findContours(combined_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        for contour in contours:
            area = cv2.contourArea(contour)
            if area < 500:  # 너무 작은 영역 제외
                continue

            # 중심점 계산
            M = cv2.moments(contour)
            if M['m00'] > 0:
                cx = int(M['m10'] / M['m00'])
                cy = int(M['m01'] / M['m00'])

                # 바운딩 박스
                x, y, w, h = cv2.boundingRect(contour)

                all_regions.append({
                    'color': color_name,
                    'center': (cx, cy),
                    'bbox': (x, y, w, h),
                    'area': area
                })

    return all_regions


def find_table_bounds_from_grid(h_coords, v_coords, img_shape, colored_regions=None):
    """그리드 라인에서 테이블 범위 찾기 (고정 25층 x 10호)"""
    if not h_coords or not v_coords:
        return None

    img_h, img_w = img_shape[:2]

    print(f"감지된 라인: 수평 {len(h_coords)}개, 수직 {len(v_coords)}개", file=sys.stderr)

    # 라인 수가 부족하면 실패
    if len(h_coords) < NUM_FLOORS + 1 or len(v_coords) < NUM_UNITS + 1:
        print(f"라인 부족: 필요 수평 {NUM_FLOORS+1}, 수직 {NUM_UNITS+1}", file=sys.stderr)
        return None

    if not colored_regions:
        selected_h = h_coords[:NUM_FLOORS + 1]
        selected_v = v_coords[:NUM_UNITS + 1]
        return (selected_v[0], selected_h[0], selected_v[-1], selected_h[-1]), selected_h, selected_v

    # 1단계: 수직선 (열) 선택 - 색상 데이터가 가장 많은 10열 구간 찾기
    best_v_start = 0
    max_v_score = -1

    for i in range(len(v_coords) - NUM_UNITS):
        v_window = v_coords[i : i + NUM_UNITS + 1]
        x_min, x_max = v_window[0], v_window[-1]

        # 해당 구간의 색상 데이터 수
        points_in_window = [r for r in colored_regions if x_min < r['center'][0] < x_max]
        data_count = len(points_in_window)

        # 간격 일관성
        gaps = [v_window[j+1] - v_window[j] for j in range(NUM_UNITS)]
        gap_std = np.std(gaps)
        gap_consistency = 1.0 / (1.0 + gap_std)

        # Y축 분포 (실제 테이블은 세로로 길게 뻗어있음)
        y_span_score = 0
        if points_in_window:
            ys = [p['center'][1] for p in points_in_window]
            y_span = max(ys) - min(ys)
            y_span_score = min(y_span / (img_h * 0.4), 1.5) * 50

        score = (data_count * 2) + (gap_consistency * 20) + y_span_score

        if score > max_v_score:
            max_v_score = score
            best_v_start = i

    selected_v = v_coords[best_v_start : best_v_start + NUM_UNITS + 1]
    x_min_final, x_max_final = selected_v[0], selected_v[-1]

    # 2단계: 수평선 (행) 선택 - 색상 데이터가 가장 많은 25행 구간 찾기
    table_data_points = [r for r in colored_regions if x_min_final < r['center'][0] < x_max_final]

    best_h_start = 0
    max_h_score = -1

    for i in range(len(h_coords) - NUM_FLOORS):
        h_window = h_coords[i : i + NUM_FLOORS + 1]
        y_min, y_max = h_window[0], h_window[-1]

        # 해당 구간의 색상 데이터 수
        data_count = sum(1 for r in table_data_points if y_min < r['center'][1] < y_max)

        # 간격 일관성
        gaps = [h_window[j+1] - h_window[j] for j in range(NUM_FLOORS)]
        gap_std = np.std(gaps)
        gap_consistency = 100.0 / (1.0 + gap_std)

        # 활성 행 수 (데이터가 있는 행)
        active_rows = 0
        for j in range(NUM_FLOORS):
            if any(h_window[j] < r['center'][1] < h_window[j+1] for r in table_data_points):
                active_rows += 1

        score = data_count + gap_consistency + (active_rows * 30)

        if score > max_h_score:
            max_h_score = score
            best_h_start = i

    selected_h = h_coords[best_h_start : best_h_start + NUM_FLOORS + 1]

    table_bounds = (selected_v[0], selected_h[0], selected_v[-1], selected_h[-1])

    print(f"선택된 그리드: {NUM_FLOORS}층 x {NUM_UNITS}호 (시작: h={best_h_start}, v={best_v_start})", file=sys.stderr)
    return table_bounds, selected_h, selected_v


def extract_cell_text(img, x1, y1, x2, y2):
    """셀에서 텍스트 추출"""
    if not OCR_AVAILABLE or OCR_READER is None:
        return ""

    # 셀 영역 추출 (여백 포함)
    margin = 2
    cell_x1 = max(0, x1 + margin)
    cell_y1 = max(0, y1 + margin)
    cell_x2 = min(img.shape[1], x2 - margin)
    cell_y2 = min(img.shape[0], y2 - margin)

    if cell_x2 <= cell_x1 or cell_y2 <= cell_y1:
        return ""

    cell_img = img[cell_y1:cell_y2, cell_x1:cell_x2]

    if cell_img.size == 0:
        return ""

    try:
        # EasyOCR로 텍스트 인식
        results = OCR_READER.readtext(cell_img, detail=0, paragraph=False)
        if results:
            # 결과 합치기 (공백 제거)
            text = ''.join(results).strip()
            # 특수문자 정리 (○, ●, □ 등 유지)
            return text
    except Exception as e:
        pass

    return ""


def map_regions_to_grid_v2(regions, h_lines, v_lines, num_floors, num_units, img=None):
    """색상 영역을 그리드 라인 기반으로 매핑 (텍스트 인식 포함)"""
    # 결과 그리드 초기화
    grid = {}
    text_grid = {}  # 텍스트 저장용
    for floor in range(1, num_floors + 1):
        grid[floor] = {}
        text_grid[floor] = {}
        for unit in range(1, num_units + 1):
            grid[floor][unit] = 'WHITE'
            text_grid[floor][unit] = ''
    
    mapped_count = 0

    for region in regions:
        cx, cy = region['center']
        color = region['color']

        # 어느 셀에 속하는지 찾기
        # 행 찾기
        row_idx = -1
        for i in range(len(h_lines) - 1):
            if h_lines[i] <= cy < h_lines[i + 1]:
                row_idx = i
                break

        # 열 찾기
        col_idx = -1
        for i in range(len(v_lines) - 1):
            if v_lines[i] <= cx < v_lines[i + 1]:
                col_idx = i
                break

        # 유효성 검사
        if 0 <= row_idx < num_floors and 0 <= col_idx < num_units:
            floor = num_floors - row_idx  # 25층~1층 (위에서 아래로)
            unit = col_idx + 1  # 1호~10호

            if grid[floor][unit] != 'WHITE':
                print(f"경고: {floor}층 {unit}호에 중복 색상 ({grid[floor][unit]} -> {color})", file=sys.stderr)

            grid[floor][unit] = color
            mapped_count += 1
        else:
            print(f"범위 벗어남: 색상={color}, 위치=({cx}, {cy}), 격자=({row_idx}, {col_idx})", file=sys.stderr)

    # 텍스트 인식 (이미지가 제공된 경우)
    if img is not None and OCR_AVAILABLE:
        print("텍스트 인식 시작...", file=sys.stderr)
        text_count = 0
        for row_idx in range(min(len(h_lines) - 1, num_floors)):
            for col_idx in range(min(len(v_lines) - 1, num_units)):
                floor = num_floors - row_idx
                unit = col_idx + 1

                # 셀 경계
                x1, x2 = v_lines[col_idx], v_lines[col_idx + 1]
                y1, y2 = h_lines[row_idx], h_lines[row_idx + 1]

                # 텍스트 추출
                text = extract_cell_text(img, x1, y1, x2, y2)
                if text:
                    text_grid[floor][unit] = text
                    text_count += 1

        print(f"텍스트 인식 완료: {text_count}개 셀", file=sys.stderr)

    print(f"매핑 완료: {mapped_count}개 성공", file=sys.stderr)

    return grid, text_grid


def process_image(image_path):
    """이미지 처리 메인 함수 (그리드 라인 기반)"""
    img = load_image(image_path)
    img_h, img_w = img.shape[:2]

    print(f"이미지 크기: {img_w} x {img_h}", file=sys.stderr)

    # 1단계: 색상 영역 검출 (그리드 검출 및 보간에 필요)
    all_regions = find_all_colored_regions(img)
    print(f"검출된 색상 영역: {len(all_regions)}개", file=sys.stderr)
    
    # 2단계: 그리드 라인 검출 (색상 영역 정보 활용)
    h_coords, v_coords = detect_table_grid(img, all_regions)
    
    # 3단계: 테이블 범위 찾기 (고정 25층 x 10호)
    result = find_table_bounds_from_grid(h_coords, v_coords, img.shape, all_regions)

    if result is None:
        print("그리드 검출 실패", file=sys.stderr)
        return None

    table_bounds, selected_h, selected_v = result
    x1, y1, x2, y2 = table_bounds

    print(f"테이블 범위: ({x1}, {y1}) ~ ({x2}, {y2})", file=sys.stderr)

    # 4단계: 격자 매핑 (텍스트 인식 포함)
    grid, text_grid = map_regions_to_grid_v2(all_regions, selected_h, selected_v, NUM_FLOORS, NUM_UNITS, img)

    # 최종 색상 분포
    final_count = {'GREEN': 0, 'YELLOW': 0, 'PINK': 0, 'WHITE': 0}
    for floor in grid:
        for unit in grid[floor]:
            final_count[grid[floor][unit]] += 1
    print(f"최종 분포: {final_count}", file=sys.stderr)

    # 디버그 이미지 생성
    debug_img = img.copy()
    cv2.rectangle(debug_img, (x1, y1), (x2, y2), (0, 255, 0), 3)

    for y in selected_h:
        cv2.line(debug_img, (x1, y), (x2, y), (0, 0, 255), 2)

    for x in selected_v:
        cv2.line(debug_img, (x, y1), (x, y2), (255, 0, 0), 2)

    for region in all_regions:
        cx, cy = region['center']
        color_bgr = {
            'GREEN': (0, 200, 0),
            'YELLOW': (0, 255, 255),
            'PINK': (255, 0, 255)
        }.get(region['color'], (255, 255, 255))
        cv2.circle(debug_img, (cx, cy), 8, color_bgr, -1)
        cv2.circle(debug_img, (cx, cy), 10, (0, 0, 0), 2)

    debug_path = os.path.join(os.path.dirname(image_path), 'debug_grid.jpg')
    cv2.imencode('.jpg', debug_img)[1].tofile(debug_path)
    print(f"디버그 이미지: {debug_path}", file=sys.stderr)

    # 결과 포맷 변환 (25층 x 10호)
    results = []
    for floor_num in range(NUM_FLOORS, 0, -1):  # 25층 → 1층
        floor_data = {
            'floor': f'{floor_num}층',
            'units': {}
        }

        for unit_num in range(1, NUM_UNITS + 1):  # 1호 → 10호
            color = grid.get(floor_num, {}).get(unit_num, 'WHITE')
            text = text_grid.get(floor_num, {}).get(unit_num, '')
            floor_data['units'][f'{unit_num}호'] = {
                'text': text,
                'color': color
            }

        results.append(floor_data)

    print(f"처리 완료: {len(results)}층", file=sys.stderr)

    return results


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "이미지 경로가 필요합니다."}))
        sys.exit(1)

    image_path = sys.argv[1]

    try:
        data = process_image(image_path)
        if data is None:
            print(json.dumps({"error": "그리드 검출 실패"}))
            sys.exit(1)
        print(json.dumps(data, ensure_ascii=False))
    except Exception as e:
        import traceback
        traceback.print_exc(file=sys.stderr)
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
