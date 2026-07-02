# -*- coding: utf-8 -*-
"""
회의록 자동 작성을 처리하는 핵심 파이썬 스크립트입니다.
"""

template_path = "../resources/minutes_template.docx" # 사용할 양식 파일 경로
sample_path = "../../../실습1-1/minutes_sample.docx" # 분석할 회의록 원본 파일 경로

print("회의록 작성을 시작합니다...") # 시작 안내 메시지
print(f"템플릿 경로: {template_path}") # 템플릿 파일 경로 출력
print(f"샘플 회의록 경로: {sample_path}") # 원본 파일 경로 출력
