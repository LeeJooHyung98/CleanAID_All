#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VB6 Get_Decrypt 함수를 Python으로 구현
basEncode.bas의 복호화 알고리즘 시뮬레이션
"""

def left_shift(s):
    """문자열 왼쪽으로 1칸 회전 (첫 문자가 맨 뒤로)"""
    if len(s) > 0:
        return s[1:] + s[0]
    return s


def right_shift(s):
    """문자열 오른쪽으로 1칸 회전 (마지막 문자가 맨 앞으로)"""
    if len(s) > 0:
        return s[-1] + s[:-1]
    return s


def scramble_wheels(wheel1, wheel2, password):
    """비밀번호로 휠 스크램블"""
    for i in range(len(password)):
        for k in range(ord(password[i]) * (i + 1)):
            wheel1 = left_shift(wheel1)
            wheel2 = right_shift(wheel2)
    return wheel1, wheel2


def get_decrypt(input_str, password):
    """VB6 Get_Decrypt 함수와 동일한 복호화"""
    # basEncode.bas의 상수값 (원본 그대로)
    WHEEL1 = 'ABCDEFGHIJKLMNOPQRSTVUWXYZ_1234567890qwertyuiopasd!@#$%^&*(),. ~`-=\\?/\'"fghjklzxcvbnm'
    WHEEL2 = 'IWEHJKTLZVOPFG_1234567890qwerBNMQRYUASDXCfghjklzxc ~`-=\\?/\'"!@#$%^&*(),.vbnmtyuiopasd'

    # 줄바꿈, 탭 제거
    input_str = input_str.replace('\n', '').replace('\r', '').replace('\t', '')

    # 비밀번호로 휠 스크램블
    wheel1, wheel2 = scramble_wheels(WHEEL1, WHEEL2, password)

    result = ''

    # 각 문자 처리
    for i, c in enumerate(input_str, 1):
        # VB의 InStr은 1-based, Python의 find는 0-based이므로 동일하게 처리
        k = wheel1.find(c)

        if k >= 0:
            result += wheel2[k]
            print(f"문자 #{i}: '{c}' → WHEEL1 위치={k+1} → WHEEL2[{k}]='{wheel2[k]}'")
        else:
            result += c
            print(f"문자 #{i}: '{c}' → WHEEL에 없음 → 그대로 '{c}'")

        # 휠 회전
        wheel1 = left_shift(wheel1)
        wheel2 = right_shift(wheel2)

    return result


if __name__ == '__main__':
    print("=" * 60)
    print("VB6 Get_Decrypt 시뮬레이션")
    print("=" * 60)

    # 테스트 케이스
    test_input = "10XX"
    test_password = ""

    print(f"\n입력값: \"{test_input}\"")
    print(f"비밀번호: \"{test_password}\"")
    print(f"\n처리 과정:\n")

    result = get_decrypt(test_input, test_password)

    print(f"\n{'=' * 60}")
    print(f"최종 결과: \"{result}\"")
    print(f"{'=' * 60}")

    # 추가 테스트
    print("\n\n[추가 테스트]")
    print("-" * 60)

    test_cases = [
        ("ABC", ""),
        ("123", ""),
        ("10XX", "MyPass"),  # 비밀번호 있는 경우
    ]

    for input_val, pwd in test_cases:
        result = get_decrypt(input_val, pwd)
        print(f"Get_Decrypt(\"{input_val}\", \"{pwd}\") = \"{result}\"")
