# -*- coding: utf-8 -*-
"""사랑
한컴오피스 한글 프로그램을 원격 제어하여 오늘의 학습 요약 문서를 생성하는 스크립트입니다.
"""

import os # OS 관련 기능을 사용하기 위한 모듈
import win32com.client as win32 # 윈도우 COM 오브젝트를 제어하여 한글을 다루기 위한 모듈

def insert_hwp_text(hwp, text):
    # 한글 OLE 액션을 활용하여 텍스트를 안전하게 입력하는 헬퍼 함수
    act = hwp.CreateAction("InsertText") # 텍스트 삽입 액션 객체 생성
    pset = act.CreateSet() # 액션 파라미터 셋 생성
    act.GetDefault(pset) # 파라미터 기본값 로드
    pset.SetItem("Text", text) # 삽입할 텍스트 데이터 대입
    act.Execute(pset) # 액션 최종 실행

def generate_hwp_summary(output_path):
    # 한글(HWP) 프로그램을 열어 요약본을 작성하고 저장하는 함수
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject") # 한글 제어 객체 생성 및 인스턴스 실행
    
    # 윈도우 창을 숨김(Visible = False) 상태로 제어하여 초점 충돌 및 대기 팝업 방지
    hwp.XHwpWindows.Item(0).Visible = False # 한글 프로그램 창 시각화 비활성화
    
    # 오늘 공부한 핵심 내용 텍스트 정의
    title_text = "★ 오늘의 바이브코딩 학습 요약 보고서 ★\r\n\r\n" # 문서의 메인 제목 텍스트
    
    body_text = (
        "1. 파이썬 개발 환경 구축\r\n"
        "  - 시스템 내에 파이썬 3.12.10 버전이 정상적으로 작동하도록 설치하였습니다.\r\n"
        "  - 파이썬 실행 경로(PATH) 등록을 완료하여 터미널 어디서든 실행 가능합니다.\r\n\r\n"
        "2. Git & GitHub 버전 관리 연동\r\n"
        "  - 로컬 학습 폴더를 Git 저장소로 초기화하고 첫 README.md 파일을 커밋하였습니다.\r\n"
        "  - 인터넷의 깃허브 원격 저장소(galua2001/vibe-coding-study)와 로컬 폴더를 연결하여 자동 업로드(Push)를 세팅하였습니다.\r\n\r\n"
        "3. uv 초고속 가상환경 시스템 세팅\r\n"
        "  - Rust 기반의 초고속 가상환경 도구인 uv를 컴퓨터에 설치하였습니다.\r\n"
        "  - 프로젝트 전용 가상환경(.venv)을 만들고, 무거운 폴더가 깃허브에 올라가지 않도록 .gitignore 설정을 마쳤습니다.\r\n\r\n"
        "4. 커스텀 에이전트 규칙(Rules) 적용\r\n"
        "  - 전역 규칙과 프로젝트 맞춤 규칙(AGENTS.md)을 작성하여 에이전트와 한글로 소통하고, 변수 옆 주석 달기, 에러 3줄 요약 규칙을 완벽하게 탑재시켰습니다.\r\n\r\n"
        "5. 회의록 자동화(Minutes_Expert) 및 문서 번역(Docx_Translator) 스킬 개발\r\n"
        "  - 파이썬 docx와 deep-translator 패키지를 가상환경에 추가로 도입했습니다.\r\n"
        "  - 설계, 개발, 검증 하위 에이전트 팀과 협업하여 회의록 자동 정돈 스킬 및 영어 문서를 한글로 자동 번역해 주는 스킬 제작에 성공하였습니다.\r\n"
    ) # 문서에 들어갈 주요 본문 텍스트
    
    # 한글 문서에 텍스트 입력하기 (헬퍼 함수 사용)
    insert_hwp_text(hwp, title_text) # 제목 텍스트 입력
    insert_hwp_text(hwp, body_text) # 본문 요약 내용 입력
    
    # 문서 저장 및 한글 프로그램 종료
    hwp.SaveAs(output_path, "HWP") # 지정된 파일 경로로 한글 HWP 문서 저장
    hwp.Quit() # 실행된 한글 프로그램 안전하게 종료
    print(f"[성공] 오늘의 학습 요약 한글(HWP) 문서가 생성되었습니다: {output_path}")

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__)) # 현재 실행 중인 스크립트 파일의 절대 경로
    hwp_output_path = os.path.abspath(os.path.join(script_dir, "오늘의_학습_요약.hwp")) # 저장할 HWP 파일 절대 경로 변환
    print("한글(HWP) 문서 생성을 시작합니다...") # 시작 안내 출력
    generate_hwp_summary(hwp_output_path) # 한글 문서 생성 함수 실행
