---
name: Docx_Translator
description: 영어 또는 다른 언어로 작성된 docx 문서의 내용을 한국어로 번역하여 새로운 한글 docx 문서를 생성하는 스킬입니다.
---

# 문서 번역 전문가 지침 (Docx Translator Guideline)

이 스킬은 원본 docx 파일의 레이아웃과 폰트 서식을 최대한 유지하면서, 내부 텍스트만 한국어(한글)로 매끄럽게 번역하여 결과물 문서를 저장하는 것을 목표로 합니다.

## 수행 단계 (Execution Steps)
1. 번역할 대상 원본 문서(예: `input.docx`)를 준비합니다.
2. `scripts/translate_docx.py` 스크립트를 실행하여 문서 내부의 문단(paragraphs) 및 표(tables) 텍스트를 읽어옵니다.
3. 외부 번역 API 또는 라이브러리(deep-translator 등)를 통해 텍스트를 한국어로 번역합니다.
4. 원본 서식을 손상시키지 않고 번역된 한국어 텍스트로 치환하여 `translated_output.docx`로 저장합니다.
