# -*- coding: utf-8 -*-
"""
win32com 한글 OLE API를 이용하여 특정 docx 문서를 hwp 문서로 일괄 변환해 주는 스크립트입니다.
"""

import os # OS 파일 경로 제어를 위한 모듈
import sys # 시스템 인자 처리를 위한 모듈
import win32com.client as win32 # 한글 원격 제어를 위한 윈도우 COM 모듈

def convert_docx_to_hwp(input_docx_path, output_hwp_path):
    # docx 워드 문서를 열어 hwp 한글 문서 파일로 변환 저장하는 함수
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject") # 한글 프로그램 실행 및 객체 생성
    
    # 윈도우 보안 모듈 작동 승인
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckDLL") # 파일 보안 경고창 우회 모듈 등록
    
    # 작업 진행 중 초점 충돌 방지를 위해 백그라운드로 실행
    hwp.XHwpWindows.Item(0).Visible = False # 한글 프로그램 화면 비활성화
    
    # 한글 API의 Open 메서드는 절대 경로만 인식하므로 절대로 변환
    abs_input_path = os.path.abspath(input_docx_path) # 원본 docx 파일의 절대 경로 변환
    abs_output_path = os.path.abspath(output_hwp_path) # 저장할 hwp 파일의 절대 경로 변환
    
    # 문서 열기 실행
    open_success = hwp.Open(abs_input_path) # 원본 워드 문서 파일 열기 실행
    
    if open_success: # 파일 열기가 정상적으로 완료된 경우
        # 다른 이름으로 저장 (포맷: HWP 문서)
        hwp.SaveAs(abs_output_path, "HWP") # 지정 경로로 한글 문서 저장 수행
        print(f"[성공] 변환이 완료되었습니다: {abs_output_path}")
    else: # 파일을 불러오지 못한 경우
        print(f"[실패] 원본 문서를 열지 못했습니다: {abs_input_path}")
        
    hwp.Quit() # 구동된 한글 프로그램 최종 종료

# 실행 영역
if __name__ == "__main__":
    # 기본 경로 설정 (실습1-1 폴더의 오늘의_학습_요약 파일 대상)
    script_dir = os.path.dirname(os.path.abspath(__file__)) # 현재 스크립트 파일의 절대 경로
    
    default_input = os.path.join(script_dir, "../../../실습1-1/오늘의_학습_요약.docx") # 기본 변환 대상 docx 경로
    default_output = os.path.join(script_dir, "../../../실습1-1/오늘의_학습_요약.hwp") # 기본 변환 결과 hwp 경로
    
    # 만약 명령줄 인자로 파일 경로가 넘어온 경우 해당 경로 사용
    input_path = sys.argv[1] if len(sys.argv) > 1 else default_input # 입력 파일 경로 변수 지정
    output_path = sys.argv[2] if len(sys.argv) > 2 else default_output # 출력 파일 경로 변수 지정
    
    print("워드 ➡️ 한글(HWP) 파일 변환 스킬 가동...") # 실행 개시 메시지
    convert_docx_to_hwp(input_path, output_path) # 변환 함수 호출
