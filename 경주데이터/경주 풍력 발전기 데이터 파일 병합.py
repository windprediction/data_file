import zipfile
import pandas as pd
from google.colab import drive
import os

# 1. 구글 드라이브 마운트
drive.mount('/content/drive')

# 2. zip 파일 경로 및 압축 해제 위치 설정
zip_path = '/content/drive/MyDrive/풍력발전/3-2. 분석용데이터_경주풍력_SCADA.zip'
extract_dir = '/content/경주풍력_SCADA'

# 3. 압축 해제
with zipfile.ZipFile(zip_path, 'r') as zip_ref:
    zip_ref.extractall(extract_dir)

# 4. 에너지 컬럼명 정의
energy_col = 'Energy Production\nActive Energy Production\n[KWh]'
full_df = pd.DataFrame()

# 5. 압축 해제된 폴더에서 연도별 파일 반복 탐색
for root, _, files in os.walk(extract_dir):
    for f in files:
        if f.endswith('.xlsx') and 'dynamic_report_ewp02_' in f:
            file_path = os.path.join(root, f)
            xls = pd.ExcelFile(file_path)

            # WTG01 ~ WTG09 시트 병합
            for sheet in xls.sheet_names:
                df = xls.parse(sheet, header=5)
                if 'Date/Time' in df.columns and energy_col in df.columns:
                    df = df[['Date/Time', energy_col]].copy()
                    df['Date/Time'] = pd.to_datetime(df['Date/Time'], errors='coerce')
                    df[energy_col] = pd.to_numeric(df[energy_col], errors='coerce')
                    full_df = pd.concat([full_df, df], ignore_index=True)

# 6. 데이터 정리
full_df.dropna(inplace=True)
full_df.sort_values('Date/Time', inplace=True)

# 7. 1시간 단위로 그룹화하여 총 발전량 계산
hourly_df = full_df.groupby(pd.Grouper(key='Date/Time', freq='1h'))[energy_col].sum().reset_index()

# 8. train_y와 비교할 수 있도록 end_datetime 기준 컬럼 추가
hourly_df['end_datetime'] = hourly_df['Date/Time'] + pd.Timedelta(hours=1)
hourly_df.drop(columns='Date/Time', inplace=True)
hourly_df.rename(columns={energy_col: 'energy_kwh'}, inplace=True)

# 9. 결과 출력
hourly_df
