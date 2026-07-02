---
name: Minutes_Expert
description: 회의록 텍스트(docx)를 읽고, 템플릿 docx 양식에 맞게 채워 넣어 최종 회의록을 자동 생성하는 스킬입니다.
---

# 회의록 작성 전문가 지침 (Minutes Expert Guideline)

이 스킬은 텍스트 기반의 회의록 내용(minutes_sample.docx)을 분석하여, 회사 공식 양식인 템플릿(minutes_template.docx)에 맞춰 변환하는 것을 목표로 합니다.

## 수행 단계 (Execution Steps)
1. `resources/minutes_template.docx` 파일을 읽어 템플릿 양식을 확인합니다.
2. `scripts/process_minutes.py` 스크립트를 사용하여 회의록 내용을 분석하고 템플릿의 변수 영역을 채웁니다.
3. 최종 생성된 결과물 파일인 `output_minutes.docx` 파일을 생성합니다.
