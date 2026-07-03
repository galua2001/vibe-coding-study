---
name: Docx_to_Hwp
description: 윈도우 한글(HWP) 프로그램을 제어하여, 특정 docx 워드 문서를 한글 문서(.hwp) 형식으로 자동 변환해 주는 스킬입니다.
---

# 워드 한글 변환 전문가 지침 (Docx to Hwp Guideline)

이 스킬은 지정된 docx 파일을 한컴오피스 한글 API를 통해 열고, 서식이나 레이아웃 유실 없이 한글 문서(*.hwp) 형식으로 저장하여 출력하는 것을 목표로 합니다.

## 수행 단계 (Execution Steps)
1. 변환할 원본 docx 파일의 경로를 확보합니다. (예: `실습1-1/오늘의_학습_요약.docx`)
2. `scripts/convert_docx_to_hwp.py` 스크립트를 실행하여 한글 프로그램을 백그라운드로 실행합니다.
3. 스크립트 내부에서 원본 docx 파일을 열고, 지정된 경로에 한글(*.hwp) 포맷으로 변환하여 저장합니다.
