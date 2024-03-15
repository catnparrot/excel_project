import os
import pathlib
import pandas as pd

# 연도 설정
YEAR = "2023"

# 점검대상
target = "pathName"

# 추출할 데이터 위치
data_locations = ['E12', 'E13', 'E19', 'E20']  # 데이터 위치 지정

# 데이터가 저장될 빈 DataFrame 생성
output_data = pd.DataFrame()

# 파일 경로 설정
folder_base_path_string = f'C:\\Users\\man\\dir\\{target}\\folder'
#folder_base_path_obj = pathlib.Path(folder_base_path_string)
#folder_base_path = os.path.dirname(folder_base_path_obj)

# 1월~12월 데이터 설정
for MONTH in range(1, 13):
    # 파일 경로 설정
    #folder_path_string = f'{folder_base_path_string}\\{YEAR}\\{target} {YEAR}년 {MONTH}월 문서'
    folder_path = f'{folder_base_path_string}\\{YEAR}\\{target} {YEAR}년 {MONTH}월 문서'
    #folder_path_obj = pathlib.Path(folder_path_string)
    #folder_path = os.path.join(folder_base_path, folder_path_obj)
    #folder_path = os.path.dirname(folder_path_obj)
    
    #print(f'folder_path: {folder_path}')
    #print(f'folder_path_obj: {folder_path_obj}')
    #print(f'folder_path_string: {folder_path_string}')

    if os.path.exists(folder_path) and os.path.isdir(folder_path):  # 디렉토리가 존재하는지 확인
        # 디렉토리 내의 파일을 검색
        for filename in os.listdir(folder_path):
            if filename.endswith('.xlsx'):  # 엑셀 파일인 경우만 처리
                filepath = os.path.join(folder_path, filename)
                
                print(f'filepath: {filepath}')

                # 엑셀 파일에서 모든 시트를 읽어옴
                excel_data = pd.read_excel(filepath, sheet_name=None, header=None)
                
                print(f'excel_data: {excel_data}')

                # 각 시트에서 데이터 위치 추출
                for sheet_name, sheet_data in excel_data.items():
                    for location in data_locations:

                        print(f'test-location: {location}')

                        row, col = pd.ExcelWriter.parse(location)
                        try:
                            value = sheet_data.iloc[row - 1, col - 1]
                        except IndexError:
                            value = None

                        # 추출한 데이터를 DataFrame에 추가
                        output_data = output_data.append(
                            {'File': filename, 'Sheet': sheet_name, 'Cell': location, 'Value': value},
                            ignore_index=True)

# 추출된 데이터를 하나의 엑셀 파일에 저장
output_file_string = f'C:\\Users\\man\\dir\\{target}\\data.xlsx'  # 결과 파일 경로로 수정
output_data.to_excel(output_file_string, index=False)

print("데이터 추출이 완료되었습니다. 결과는 '{}'에 저장되었습니다.".format(output_file_string))
