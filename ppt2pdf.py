import os
import glob
import win32com.client
import time

# --- 설정 ---
INPUT_DIR = r"C:\Users\lsh\Desktop\강의자료"   # PPT 파일이 있는 폴더
OUTPUT_DIR = r"C:\Users\lsh\Desktop\강의자료" # PDF 파일을 저장할 폴더
# ---------------

# 절대 경로로 변환 (COM 객체가 상대 경로를 잘 처리 못 할 수 있음)
input_path = os.path.abspath(INPUT_DIR)
output_path = os.path.abspath(OUTPUT_DIR)

# 입출력 폴더 생성
os.makedirs(input_path, exist_ok=True)
os.makedirs(output_path, exist_ok=True)

# PowerPoint 애플리케이션 실행
try:
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1 # 1 = 보이게, 0 = 안 보이게
except Exception as e:
    print(f"PowerPoint 실행 실패: {e}")
    print("PowerPoint가 설치되어 있는지 확인하세요.")
    exit()

# input_ppt 폴더에서 .ppt 및 .pptx 파일 모두 찾기
ppt_files = glob.glob(os.path.join(input_path, "*.ppt*"))

if not ppt_files:
    print(f"'{INPUT_DIR}' 폴더에 변환할 PPT 파일이 없습니다.")
    powerpoint.Quit()
    exit()
    
print(f"총 {len(ppt_files)}개의 파일을 변환합니다...")

# PDF 포맷 코드 (PowerPoint 상수)
ppFormatPDF = 32

for ppt_file in ppt_files:
    print(f"변환 중: {ppt_file}")
    
    try:
        # 프레젠테이션 열기
        presentation = powerpoint.Presentations.Open(ppt_file, WithWindow=False)
        
        # 파일 이름 설정 (확장자 변경)
        file_name = os.path.basename(ppt_file)
        pdf_name = os.path.splitext(file_name)[0] + ".pdf"
        pdf_path = os.path.join(output_path, pdf_name)
        
        # PDF로 저장
        presentation.SaveAs(pdf_path, ppFormatPDF)
        
        print(f"변환 완료: {pdf_name}")
        
        # 프레젠테이션 닫기
        presentation.Close()
        
    except Exception as e:
        print(f"'{ppt_file}' 변환 중 오류 발생: {e}")

# PowerPoint 종료
powerpoint.Quit()
print("모든 작업이 완료되었습니다.")
