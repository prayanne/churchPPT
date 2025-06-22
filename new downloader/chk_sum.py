import os
import sys
import re

def extract_number(filename):
    match = re.match(r"\[새찬송가\]\s*(\d+)장\.pptx$", filename)
    return int(match.group(1)) if match else None

def check_missing_numbers(dir_path):
    if not os.path.isdir(dir_path):
        print(f"'{dir_path}'는 유효한 디렉토리가 아닙니다.")
        return

    numbers = []
    for file in os.listdir(dir_path):
        number = extract_number(file)
        if number is not None:
            numbers.append(number)

    if not numbers:
        print("숫자가 포함된 '[새찬송가] ***장.pptx' 파일이 없습니다.")
        return

    numbers = sorted(set(numbers))
    print(f"감지된 번호: {numbers}")

    full_range = set(range(numbers[0], numbers[-1] + 1))
    missing = sorted(full_range - set(numbers))

    if missing:
        print(f"❌ 빠진 숫자: {missing}")
    else:
        print("✅ 모든 번호가 연속적으로 존재합니다.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("사용법: python check_missing_hymns.py [디렉토리_경로]")
    else:
        check_missing_numbers(sys.argv[1])
