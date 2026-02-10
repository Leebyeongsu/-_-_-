import sys
import json
import cv2
import numpy as np
import os

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# EasyOCR 초기화
try:
    import easyocr
    OCR_READER = easyocr.Reader(['en', 'ko'], gpu=False, verbose=False)
    OCR_AVAILABLE = True
    print("EasyOCR 로드 완료", file=sys.stderr)
except:
    OCR_READER = None
    OCR_AVAILABLE = False


def load_image(image_path):
    """이미지 로드"""
    img_array = np.fromfile(image_path, np.uint8)
    img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    if img is None:
        raise Exception(f"이미지 로드 실패: {image_path}")
    return img


def find_grid_lines(img):
    """모든 그리드 라인 찾기"""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape

    # 이진화
    _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

    # 수평선 감지
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (w // 10, 1))
    h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel)

    # 수직선 감지
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, h // 15))
    v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel)

    # 수평선 y좌표 추출
    h_coords = []
    for y in range(h):
        if np.sum(h_lines[y, :]) > w * 50:
            h_coords.append(y)

    # 수직선 x좌표 추출
    v_coords = []
    for x in range(w):
        if np.sum(v_lines[:, x]) > h * 30:
            v_coords.append(x)

    # 근접 라인 병합
    def merge(coords, gap=5):
        if not coords:
            return []
        coords = sorted(set(coords))
        result = [coords[0]]
        for c in coords[1:]:
            if c - result[-1] > gap:
                result.append(c)
        return result

    return merge(h_coords), merge(v_coords)


def find_main_table(h_lines, v_lines, img_shape):
    """
    메인 데이터 테이블 찾기
    - 가장 많은 라인이 일정한 간격으로 배열된 영역 찾기
    - 헤더 영역 자동 제외
    """
    h, w = img_shape[:2]

    if len(h_lines) < 3 or len(v_lines) < 3:
        return None, None

    # 수평선 간격 분석
    h_gaps = [h_lines[i+1] - h_lines[i] for i in range(len(h_lines)-1)]

    # 가장 흔한 간격 찾기 (데이터 행의 높이)
    if not h_gaps:
        return None, None

    # 간격별 빈도 계산
    gap_counts = {}
    for gap in h_gaps:
        # 비슷한 간격끼리 그룹화 (±3 픽셀)
        found = False
        for key in gap_counts:
            if abs(key - gap) <= 3:
                gap_counts[key] += 1
                found = True
                break
        if not found:
            gap_counts[gap] = 1

    # 가장 많이 나타나는 간격 = 데이터 행 높이
    common_h_gap = max(gap_counts, key=gap_counts.get)
    print(f"데이터 행 높이: {common_h_gap}px (빈도: {gap_counts[common_h_gap]})", file=sys.stderr)

    # 해당 간격을 가진 연속 구간 찾기
    data_h_start = 0
    data_h_end = len(h_lines) - 1
    max_consecutive = 0
    current_start = 0
    current_count = 0

    for i in range(len(h_gaps)):
        if abs(h_gaps[i] - common_h_gap) <= 5:  # 허용 오차
            if current_count == 0:
                current_start = i
            current_count += 1
        else:
            if current_count > max_consecutive:
                max_consecutive = current_count
                data_h_start = current_start
                data_h_end = current_start + current_count
            current_count = 0

    # 마지막 구간 체크
    if current_count > max_consecutive:
        data_h_start = current_start
        data_h_end = current_start + current_count

    # 수직선도 동일하게 처리
    v_gaps = [v_lines[i+1] - v_lines[i] for i in range(len(v_lines)-1)]

    gap_counts = {}
    for gap in v_gaps:
        found = False
        for key in gap_counts:
            if abs(key - gap) <= 5:
                gap_counts[key] += 1
                found = True
                break
        if not found:
            gap_counts[gap] = 1

    common_v_gap = max(gap_counts, key=gap_counts.get)
    print(f"데이터 열 너비: {common_v_gap}px (빈도: {gap_counts[common_v_gap]})", file=sys.stderr)

    # 해당 간격을 가진 연속 구간 찾기
    data_v_start = 0
    data_v_end = len(v_lines) - 1
    max_consecutive = 0
    current_start = 0
    current_count = 0

    for i in range(len(v_gaps)):
        if abs(v_gaps[i] - common_v_gap) <= 8:
            if current_count == 0:
                current_start = i
            current_count += 1
        else:
            if current_count > max_consecutive:
                max_consecutive = current_count
                data_v_start = current_start
                data_v_end = current_start + current_count
            current_count = 0

    if current_count > max_consecutive:
        data_v_start = current_start
        data_v_end = current_start + current_count

    # 열이 한 칸씩 밀린 문제 해결: data_v_start를 0으로 조정
    # data_v_start=1일 때 data_v[0] = v_lines[1]이 실제로는 2호의 왼쪽 경계
    # data_v_start=0일 때 data_v[0] = v_lines[0]이 층 열 왼쪽, data_v[1] = v_lines[1]이 1호 왼쪽
    # 따라서 data_v_start를 0으로 조정하거나, data_v_start를 -1로 조정
    # 하지만 data_v_start는 인덱스이므로 음수가 될 수 없음
    # 대신 data_v_start를 0으로 강제 설정하거나, selected_v를 한 칸 앞으로 이동
    
    # 해결: data_v_start를 0으로 조정 (층 열 포함)
    if data_v_start > 0:
        data_v_start = data_v_start - 1  # 한 칸 앞으로 이동

    # 선택된 라인들
    # 헤더 행들 제외:
    # - 열 헤더 행 (호/층, 1, 2, 3...)
    # - 통계 값 행 (189, 35%, 66...)
    # 층 열 제외:
    # - 첫 번째 열이 층 번호 (25, 24, 23...)
    # skip 없이 전체 선택 (process_image에서 오프셋 처리)
    selected_h = h_lines[data_h_start:data_h_end+2]
    selected_v = v_lines[data_v_start:data_v_end+2]

    num_rows = len(selected_h) - 1
    num_cols = len(selected_v) - 1

    print(f"메인 테이블: {num_rows}행 x {num_cols}열", file=sys.stderr)
    print(f"수평선 범위: {data_h_start} ~ {data_h_end+1}", file=sys.stderr)
    print(f"수직선 범위: {data_v_start} ~ {data_v_end+1}", file=sys.stderr)

    return selected_h, selected_v


def classify_color(r, g, b):
    """RGB + HSV 색상 분류 - 파스텔 톤 최적화"""
    # 채널 차이 계산
    rg_diff = abs(r - g)
    rb_diff = abs(r - b)
    gb_diff = abs(g - b)
    brightness = (r + g + b) / 3

    # 1. RGB 기반 분류
    
    # 흰색 (모든 채널이 매우 높고 비슷함)
    if r > 245 and g > 245 and b > 245:
        return "WHITE"

    # 거의 흰색 (밝고 채널 차이 매우 작음)
    if brightness > 240 and max(rg_diff, rb_diff, gb_diff) < 15:
        return "WHITE"

    # 노란색 RGB 조건 1: 원본 엔진 방식 (더 엄격)
    if r > 220 and g > 220 and b < 230:
        if r > b + 10 and g > b + 10:
            if abs(r - g) < 30:  # R과 G가 비슷
                return "YELLOW"

    # 노란색 RGB 조건 2: 더 연한 톤
    if r > 200 and g > 200:
        rg_avg = (r + g) / 2
        if rg_avg > b + 5 and rg_diff < 35:
            return "YELLOW"

    # 노란색 RGB 조건 3: 매우 연한 크림/베이지 톤
    if r > 230 and g > 230:
        if r >= b and g >= b:
            if (r + g) > (b * 2 + 20):
                return "YELLOW"

    # 녹색 (G가 가장 높음)
    if g > 200:
        if g > r + 5 and g > b + 5:
            return "GREEN"

    # 분홍색 (R이 높고, B도 상대적으로 높음, G가 낮음)
    if r > 220 and b > 200:
        if r > g and b > g - 20:
            if r >= b - 30:  # R과 B가 비슷하거나 R이 더 높음
                return "PINK"

    # 2. HSV 기반 추가 판정 (원본 엔진 방식)
    r_n, g_n, b_n = r / 255.0, g / 255.0, b / 255.0
    max_c = max(r_n, g_n, b_n)
    min_c = min(r_n, g_n, b_n)
    diff = max_c - min_c

    v = max_c
    s = 0 if max_c == 0 else diff / max_c

    # HSV 계산
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
        # 노란색 영역: 40-70도 (원본 엔진 범위)
        if 40 <= h <= 70:
            return "YELLOW"

        # 녹색 영역: 70-160도
        if 70 < h <= 160:
            return "GREEN"

        # 분홍/빨강 영역: 300-360 또는 0-35도
        if h > 300 or h < 35:
            return "PINK"

        # 보라/분홍 영역: 260-300도
        if 260 <= h <= 300:
            return "PINK"

    # 밝은 회색/흰색 계열
    if brightness > 235 and max(rg_diff, rb_diff, gb_diff) < 20:
        return "WHITE"

    return "WHITE"


def detect_symbols(cell_img):
    """셀 이미지에서 기호 감지: ◎ □ ● ○"""
    if cell_img.size == 0:
        return []
    
    symbols = []
    gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape
    
    if h < 10 or w < 10:  # 너무 작은 셀
        return []
    
    # 적응형 이진화로 더 정확한 윤곽선 추출
    # THRESH_BINARY_INV: 어두운 기호(●○◎□)가 흰색으로 반전되어 윤곽선 감지 가능
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                    cv2.THRESH_BINARY_INV, 11, 2)
    
    # 윤곽선 찾기
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    detected_symbols = []  # (기호, x좌표, 면적) 튜플 리스트
    
    for contour in contours:
        area = cv2.contourArea(contour)
        if area < 20:  # 너무 작은 노이즈 제외
            continue

        # 원형도 계산
        perimeter = cv2.arcLength(contour, True)
        if perimeter == 0:
            continue
        circularity = 4 * np.pi * area / (perimeter * perimeter)

        # 중심점 계산
        M = cv2.moments(contour)
        if M["m00"] == 0:
            continue
        cx = int(M["m10"] / M["m00"])
        cy = int(M["m01"] / M["m00"])

        # 바운딩 박스
        x, y, w_rect, h_rect = cv2.boundingRect(contour)

        # 원형 기호 (◎, ○, ●) 감지
        if circularity > 0.7 and area > 30:
            # ROI에서 원본 gray 이미지의 평균 밝기 계산
            roi = gray[max(0, y):min(h, y+h_rect), max(0, x):min(w, x+w_rect)]
            if roi.size == 0:
                continue
            avg_brightness = np.mean(roi)

            # 이중원 ◎ 감지: 중심부와 외곽의 밝기 차이 확인
            is_double_circle = False
            if area > 100:  # 큰 원만 이중원 후보
                # 중심부 영역 (반지름의 40%)
                center_r = int(min(w_rect, h_rect) * 0.2)
                if center_r > 2:
                    center_roi = gray[max(0, cy-center_r):min(h, cy+center_r),
                                     max(0, cx-center_r):min(w, cx+center_r)]
                    if center_roi.size > 0:
                        center_brightness = np.mean(center_roi)
                        # 중심부가 외곽보다 훨씬 어두우면 이중원
                        if center_brightness < avg_brightness * 0.6 and center_brightness < 120:
                            is_double_circle = True

            if is_double_circle:
                detected_symbols.append(('◎', cx, area))
            elif avg_brightness < 80:  # 매우 어두운 원 → ●
                detected_symbols.append(('●', cx, area))
            elif avg_brightness > 150:  # 밝은 원 → ○
                detected_symbols.append(('○', cx, area))
            else:
                # 중간 밝기: 크기로 구분
                if area > 100:
                    detected_symbols.append(('◎', cx, area))
                elif avg_brightness < 130:
                    detected_symbols.append(('●', cx, area))
                else:
                    detected_symbols.append(('○', cx, area))

        # 사각형 □ 감지
        elif area > 40:
            # 사각형 근사
            peri = cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, 0.04 * peri, True)

            # 4개 꼭짓점이면 사각형
            if len(approx) == 4:
                aspect_ratio = float(w_rect) / h_rect if h_rect > 0 else 0
                # 정사각형에 가까움
                if 0.6 < aspect_ratio < 1.4 and area > 50:
                    detected_symbols.append(('□', cx, area))
    
    # 중복 제거 (같은 위치의 기호는 면적이 큰 것만 선택)
    if detected_symbols:
        # x 좌표로 정렬
        detected_symbols.sort(key=lambda x: x[1])

        # 중복 제거 (비슷한 x 좌표의 기호는 면적이 큰 것만)
        filtered = []
        for i, (sym, x, area) in enumerate(detected_symbols):
            is_duplicate = False
            for j, (prev_sym, prev_x, prev_area) in enumerate(filtered):
                if abs(x - prev_x) < 5:  # 5픽셀 이내면 중복
                    if area > prev_area:
                        filtered[j] = (sym, x, area)
                    is_duplicate = True
                    break
            if not is_duplicate:
                filtered.append((sym, x, area))

        # x 좌표로 다시 정렬 후 원본 기호 그대로 사용 (변환하지 않음)
        filtered.sort(key=lambda x: x[1])
        seen = set()
        for s in filtered:
            symbol = s[0]  # 원본 기호 (◎, ○, ●, □)
            if symbol and symbol not in seen:
                seen.add(symbol)
                symbols.append(symbol)

    return symbols


def extract_text(img, x1, y1, x2, y2):
    """셀에서 텍스트 및 기호 추출"""
    margin = 2
    cell = img[y1+margin:y2-margin, x1+margin:x2-margin]

    if cell.size == 0:
        return ""

    result_parts = []
    
    # 1. 기호 감지 (◎ □ ● ○)
    symbols = detect_symbols(cell)
    result_parts.extend(symbols)
    
    # 2. OCR로 문자 인식 (M, V, P 등)
    if OCR_AVAILABLE:
        try:
            # detail=1로 설정하여 위치 정보도 가져오기
            ocr_results = OCR_READER.readtext(cell, detail=1, paragraph=False)
            if ocr_results:
                # OCR 결과를 x 좌표로 정렬
                ocr_chars = []
                for (bbox, text, confidence) in ocr_results:
                    if confidence > 0.3:  # 신뢰도가 낮은 결과 제외
                        # bbox의 중심 x 좌표 계산
                        x_coords = [point[0] for point in bbox]
                        cx = sum(x_coords) / len(x_coords)
                        # 텍스트에서 유효한 문자만 추출
                        for char in text.upper():
                            if char.isalpha() and char in ['M', 'V', 'P', 'I', 'O']:
                                ocr_chars.append((char, cx))
                
                # x 좌표로 정렬
                ocr_chars.sort(key=lambda x: x[1])
                result_parts.extend([char for char, _ in ocr_chars])
        except Exception as e:
            # detail=1이 실패하면 detail=0으로 재시도
            try:
                ocr_results = OCR_READER.readtext(cell, detail=0, paragraph=False)
                if ocr_results:
                    ocr_text = ''.join(ocr_results).strip().upper()
                    for char in ocr_text:
                        if char.isalpha() and char in ['M', 'V', 'P', 'I', 'O']:
                            result_parts.append(char)
            except:
                pass
    
    # 결과 합치기 (공백 없이, 왼쪽부터 순서대로, 중복 문자 제거)
    seen = set()
    deduped = []
    for char in result_parts:
        if char and char not in seen:
            seen.add(char)
            deduped.append(char)
    return ''.join(deduped)


def extract_header_info(img, table_top_y):
    """이미지 상단 헤더 영역에서 동 번호, 아파트 이름 등 추출"""
    header_info = {
        "building": "",   # 동 번호 (예: 102동, 106동)
        "name": "",       # 아파트 이름 (예: LG신주례1차)
    }

    if not OCR_AVAILABLE or table_top_y < 20:
        return header_info

    h, w = img.shape[:2]
    # 테이블 상단 위의 영역을 헤더로 간주
    header_region = img[0:table_top_y, 0:w]

    if header_region.size == 0:
        return header_info

    try:
        results = OCR_READER.readtext(header_region, detail=1, paragraph=False)
        print(f"헤더 OCR 결과: {len(results)}개", file=sys.stderr)

        texts = []
        for (bbox, text, conf) in results:
            if conf > 0.3:
                # bbox 중심 x 좌표
                cx = sum(p[0] for p in bbox) / len(bbox)
                texts.append((cx, text.strip()))
                print(f"  헤더 텍스트: '{text.strip()}' (신뢰도: {conf:.2f}, x={cx:.0f})", file=sys.stderr)

        # 텍스트에서 정보 추출
        for _, text in texts:
            # 동 번호 찾기: 숫자 + "동" 또는 "동선택" 다음 숫자
            import re
            dong_match = re.search(r'(\d{2,4})\s*동?', text)
            if dong_match and not header_info["building"]:
                num = dong_match.group(1)
                if 100 <= int(num) <= 9999:
                    header_info["building"] = f"{num}동"

            # 아파트 이름 찾기: 한글 2글자 이상 + 숫자(차) 패턴
            if len(text) >= 4 and any(c >= '가' and c <= '힣' for c in text):
                # "동선택", "세대수" 등 메타 텍스트 제외
                skip_words = ['동선택', '세대수', '호', '층']
                if not any(sw in text for sw in skip_words):
                    if not header_info["name"]:
                        header_info["name"] = text

    except Exception as e:
        print(f"헤더 OCR 오류: {e}", file=sys.stderr)

    print(f"헤더 정보: {header_info}", file=sys.stderr)
    return header_info


def process_image(image_path):
    """이미지 처리"""
    img = load_image(image_path)
    h, w = img.shape[:2]
    print(f"이미지: {w} x {h}", file=sys.stderr)

    # 1. 그리드 라인 찾기
    h_lines, v_lines = find_grid_lines(img)
    print(f"전체 라인: 수평 {len(h_lines)}, 수직 {len(v_lines)}", file=sys.stderr)

    # 2. 메인 데이터 테이블 찾기
    data_h, data_v = find_main_table(h_lines, v_lines, img.shape)

    if data_h is None or len(data_h) < 2 or len(data_v) < 2:
        print("테이블 감지 실패", file=sys.stderr)
        return None

    num_rows = len(data_h) - 1
    print(f"감지된 행: {num_rows}", file=sys.stderr)

    # ======================================================
    # 헤더/데이터 판별: 첫 행의 첫 열(층 번호 열)을 OCR
    # ======================================================
    # 핵심 로직:
    #   - OCR 결과가 숫자(예: "25") → 첫 행은 데이터 행 (25층)
    #     → 행을 건너뛰지 않고, actual_rows = 해당 숫자로 제한
    #     → 아래쪽 여분 행은 자연스럽게 무시됨
    #   - OCR 결과가 비숫자(예: "호", "층") → 첫 행은 헤더
    #     → 첫 행을 건너뜀
    # ======================================================
    rows_to_skip = 0
    detected_max_floor = None

    if len(data_h) > 2 and len(data_v) > 1:
        margin = 3
        y1, y2 = data_h[0], data_h[1]
        x1, x2 = data_v[0], data_v[1]
        floor_cell = img[y1+margin:y2-margin, x1+margin:x2-margin]

        if floor_cell.size > 0:
            avg = cv2.mean(floor_cell)[:3]
            print(f"첫 행 첫 열: RGB=({avg[2]:.0f},{avg[1]:.0f},{avg[0]:.0f})", file=sys.stderr)

            if OCR_AVAILABLE:
                try:
                    result = OCR_READER.readtext(floor_cell, detail=0, paragraph=False)
                    text = ''.join(result).strip().replace(' ', '')
                    print(f"첫 행 첫 열 OCR: '{text}'", file=sys.stderr)

                    if text.isdigit():
                        detected_max_floor = int(text)
                        # 숫자 → 첫 행은 데이터 (층 번호)
                        # 건너뛰지 않음! actual_rows로 제한하여 여분 행 무시
                        print(f"첫 행 = {detected_max_floor}층 데이터 (건너뛰지 않음)", file=sys.stderr)
                    else:
                        # 비숫자 → 헤더 행 (호/층 등)
                        rows_to_skip = 1
                        print(f"첫 행 = 헤더 ('{text}'), 건너뜀", file=sys.stderr)

                        # 두 번째 행에서 최대 층 번호 확인
                        if len(data_h) > 3:
                            y1b, y2b = data_h[1], data_h[2]
                            cell2 = img[y1b+margin:y2b-margin, x1+margin:x2-margin]
                            if cell2.size > 0:
                                result2 = OCR_READER.readtext(cell2, detail=0, paragraph=False)
                                text2 = ''.join(result2).strip().replace(' ', '')
                                if text2.isdigit():
                                    detected_max_floor = int(text2)
                                    print(f"두 번째 행 = {detected_max_floor}층", file=sys.stderr)
                except Exception as e:
                    print(f"OCR 실패: {e}", file=sys.stderr)

    if rows_to_skip > 0:
        data_h = data_h[rows_to_skip:]
        num_rows = len(data_h) - 1
        print(f"헤더 {rows_to_skip}행 건너뜀 → {num_rows}행 남음", file=sys.stderr)

    if len(data_v) > 1:
        data_v_lines = data_v[1:]  # 층 열 제외, 1호부터 시작
    else:
        data_v_lines = data_v

    num_cols = len(data_v_lines) - 1 if len(data_v_lines) > 1 else 16

    # actual_rows: detected_max_floor가 있으면 그 값 사용 (여분 행 자동 무시)
    if detected_max_floor and detected_max_floor <= num_rows:
        actual_rows = detected_max_floor
        print(f"층 번호 기반 행 수 결정: {actual_rows}행 (감지행={num_rows}, 최대층={detected_max_floor})", file=sys.stderr)
    else:
        actual_rows = min(num_rows, 35)
        print(f"기본 행 수 결정: {actual_rows}행", file=sys.stderr)

    actual_cols = min(num_cols, 16)

    print(f"수직선: {len(data_v)}개 (전체), {len(data_v_lines)}개 (데이터 열)", file=sys.stderr)
    if len(data_v) > 0:
        print(f"data_v[0] 위치: {data_v[0]}px (층 열 경계)", file=sys.stderr)
    if len(data_v) > 1:
        print(f"data_v[1] 위치: {data_v[1]}px (1호 왼쪽 경계 예상)", file=sys.stderr)
    if len(data_v_lines) > 0:
        print(f"data_v_lines[0] 위치: {data_v_lines[0]}px (1호 왼쪽 경계)", file=sys.stderr)
    if len(data_v_lines) > 1:
        print(f"data_v_lines[1] 위치: {data_v_lines[1]}px (2호 왼쪽 경계)", file=sys.stderr)
    if len(data_v_lines) > 9:
        print(f"data_v_lines[9] 위치: {data_v_lines[9]}px (10호 왼쪽 경계)", file=sys.stderr)
    if len(data_v_lines) > 10:
        print(f"data_v_lines[10] 위치: {data_v_lines[10]}px (10호 오른쪽 경계)", file=sys.stderr)
    print(f"데이터 열: {actual_cols}개 (1호~{actual_cols}호)", file=sys.stderr)
    print(f"오프셋 적용: {actual_rows}행 x {actual_cols}열", file=sys.stderr)

    # 3. 각 셀 처리
    results = []

    for row in range(actual_rows):
        floor_num = actual_rows - row  # 25층~1층
        floor_data = {
            "floor": f"{floor_num}층",
            "units": {}
        }

        for col in range(actual_cols):
            unit_num = col + 1  # 1호~10호 (원본 엔진 방식)

            # 행 경계
            if row < len(data_h) - 1:
                y1, y2 = data_h[row], data_h[row + 1]
            else:
                continue

            # 열 경계 (원본 엔진 방식: data_v_lines[col]이 col+1호의 왼쪽 경계)
            if col < len(data_v_lines) - 1:
                x1, x2 = data_v_lines[col], data_v_lines[col + 1]
            else:
                continue

            # 색상 샘플링 (셀 전체 영역 사용, 테두리 제외)
            margin = 3  # 테두리 여백
            sy1 = max(0, y1 + margin)
            sy2 = min(h, y2 - margin)
            sx1 = max(0, x1 + margin)
            sx2 = min(w, x2 - margin)

            roi = img[sy1:sy2, sx1:sx2]
            if roi.size > 0:
                # 셀 전체 영역의 평균 색상 계산
                avg = cv2.mean(roi)[:3]
                color = classify_color(avg[2], avg[1], avg[0])
            else:
                color = "WHITE"

            # 텍스트 추출
            text = extract_text(img, x1, y1, x2, y2)

            floor_data["units"][f"{unit_num}호"] = {
                "text": text,
                "color": color
            }

        results.append(floor_data)

    # 통계
    counts = {"GREEN": 0, "YELLOW": 0, "PINK": 0, "WHITE": 0}
    for floor in results:
        for unit in floor["units"].values():
            counts[unit["color"]] += 1

    print(f"색상 분포: {counts}", file=sys.stderr)
    print(f"완료: {len(results)}층 x {actual_cols}호", file=sys.stderr)

    # 이미지 상단 헤더 정보 추출
    table_top_y = data_h[0] if len(data_h) > 0 else 0
    header_info = extract_header_info(img, table_top_y)

    return {
        "header": header_info,
        "data": results
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "이미지 경로 필요"}))
        sys.exit(1)

    try:
        result = process_image(sys.argv[1])
        if result is None:
            print(json.dumps({"error": "테이블 감지 실패"}))
            sys.exit(1)
        print(json.dumps(result, ensure_ascii=False))
    except Exception as e:
        import traceback
        traceback.print_exc(file=sys.stderr)
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
