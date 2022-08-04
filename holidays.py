'''공공데이터 포털로부터 휴일 날짜을 얻어서 datetime(YYYY MM DD 정보만) 리스트로 반환
'''

import datetime
import requests
import json
import pandas as pnds
from pandas import json_normalize

def get_dates() -> pnds.Series:
    """한국전문연구원 공공데이터에서 requests를 이용해서 접속하고, 응답으로 받은 데이터를 JSON객체로 만들어서 실제 데이터를 추출한다. 그리고
    이를 Dataframe으로 전환하여 'locdate'컬럼에서 년월일(YYYY MM DD) datetime 리스트를 반환한다

    Returns:
        Pandas.Series: 휴일 datetime Series
    """
    
    # 오늘 날짜와 올해 년도 
    today = datetime.datetime.today().strftime('%Y%m%d')                # 오늘이 휴일인지 비교할때 사용
    today_year = datetime.datetime.today().year                         # 공공데이터 접속 시 필요한 년도
    
    # 한국천문연구원_특일 정보에 사용되는 key and address
    key = 'QmNdKSuwR11m4wrVojTx8ugJVgmgJ%2F6wgBPjeDh8zVtsPfbCmE0qaiuzcQn%2B2lwqP1AwzsiQatq9ir59nO0rWg%3D%3D'
    url = 'http://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo?_type=json&numOfRows=50&solYear=' + str(today_year) + '&ServiceKey=' + str(key)
    
    # requests로 접근
    response = requests.get(url)
    
    # 응답코드 'Ok'일 때 정보 추출
    if response.status_code == 200:
        json_obj = json.loads(response.text)                            # JSON 객체로 로딩
        holidays_data = json_obj['response']['body']['items']['item']   # JSON 객체로부터 휴일 데이터 추출
        holidays_df = json_normalize(holidays_data)                     # Json 객체를 Dataframe 객체로 변환
    
    # datetime 타입으로 변경
    holidays_df['locdate'] = pnds.to_datetime(holidays_df['locdate'], format='%Y%m%d')
    
    # 휴일 datefime(YYYY MM DD) 리스트 반환
    return [day.date() for day in holidays_df['locdate']]


if __name__ == '__main__':
    print(get_dates())