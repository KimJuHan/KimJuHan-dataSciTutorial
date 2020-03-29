import numpy as np
import pandas as pd

### EXCEL 고급필터에 해당하는 문제 ###

# 데이터 가져오기
ws = pd.read_excel('C:/Users/kjy/Desktop/컴활/03-2010버전 출제유형분석3회/실습파일.xlsm', sheet_name=1)

# dtypes : datetime일 때, 연,월,일,시,분 파싱하기 -> dt 프로퍼티 year, month, day, hour, minute
# dtypes : datetime일 때, 람다 함수 쓰면 unsubscriptable 에러 난다. 당연 왜냐면 dtype이 문자열이 아니므로 기본적인 리스트 파싱 불가
# astype(str)로 타입변경을 해주면 가능
print(ws["등록일자"].astype(str).apply(lambda x : int(x[:4]) >= 2006 and int(x[:4]) <= 2008))

# multiple condition 걸어줄 때는 괄호를 꼭 해준다.
print(ws[(ws["등록일자"].dt.year >= 2006) & (ws["등록일자"].dt.year <= 2008)])
print(ws[(ws["접수코드"].apply(lambda x : int(x[1:]) < 100))])
selected = ws[(ws["접수코드"].apply(lambda x : int(x[1:]) < 100)) |
         ((ws["등록일자"].dt.year >= 2006) &
         (ws["등록일자"].dt.year <= 2008))]
print(selected)
print(selected.shape)

### EXCEL 계산작업 ###

# 원본 데이터 가져오기
# header : column label 지정, index_col : row label 지정
ws = pd.read_excel('C:/Users/kjy/Desktop/컴활/03-2010버전 출제유형분석3회/실습파일.xlsm', sheet_name=2, header=1)

# # # 데이터가 되는 부분 파싱
data = ws.iloc[:30, 1:10]
print(data)

# 데이터 value 배열
# items() 메소드 이용
values = []
for idx, value in data["과목"].items():
    if value not in values:
        values.append(value)
print(values)

# 평균을 담을 배열
# 딕셔너리 초기화
averages = { x : 0 for x in values }
print(averages)

for i in values:
    target = data[data["과목"] == i]["점수"].astype('float64')
    averages[i] = target.sum() / target.count()
print(averages)

criteria_table = ws.iloc[42:, 1:5]
criteria_table.columns = ["과목", 1, 11, 21]
print(criteria_table)
values = [ x for x in criteria_table["과목"] ]
print(values)

# 지금처럼 정렬된 데이터가 아니라 비정렬 데이터라고 가정
# data.loc[] = val로 해야지 원본 데이터가 수정됨
data = data.set_index("접수번호")
print(data)
for i in data.index:
    # 과목에 맞는 행 가져오기
    cmp = criteria_table[criteria_table["과목"] == data.loc[i, "과목"]]
    # 접수번호 비교
    if i >= 1:
        data.loc[i, "할인액"] = cmp.iloc[0, 1] * data.loc[i, "수강료"]
    elif i >= 11:
        data.loc[i, "할인액"] = cmp.iloc[0, 11] * data.loc[i, "수강료"]
    elif i >= 21 and i <= 30:
        data.loc[i, "할인액"] = cmp.iloc[0, 21] * data.loc[i, "수강료"]

print(data)

# lookup 함수 활용
# column이 숫자이므로 이를 맞춰주기 위해서는 lookup table을 duplicate 해야한다.
# 어떤 경우에 사용하냐? 특정 값으로 채워진 2차 행렬의 테이블이 있고, 그 행렬 레이블을 value값으로 가지는 또다른 테이블이 있을 경우 => 엑셀의 hlookup
criteria_table = criteria_table.set_index("과목")
tmp1 = pd.DataFrame(np.tile(criteria_table[[1]], (1, 10)))
tmp2 = pd.DataFrame(np.tile(criteria_table[[11]], (1, 10)))
tmp3 = pd.DataFrame(np.tile(criteria_table[[21]], (1, 10)))

new_criteria_table = pd.concat([tmp1, tmp2, tmp3], axis=1)
new_criteria_table.index = ["국어", "국사", "영어", "수학"]
new_criteria_table.columns = range(1, 31)

print(new_criteria_table)

print(data)
data["할인액"] = new_criteria_table.lookup(data["과목"], data["접수번호"]) * data["수강료"]
print(data)

values = {}
for i, k in data["접수코드"].items():
    if k not in values:
        values[k] = 1
    else:
        values[k] += 1

# 딕셔너리의 키값이 칼럼 레이블이 된다. 따라서 value값이 스칼라일 경우, 에러가 날 수 있다.
df = pd.DataFrame(values, index=["인원수"])
df = df.T
# sorting by index
df.sort_index(inplace=True)
# astype, apply는 모두 in_place 메소드가 아니다.
df["인원수"] = df["인원수"].astype(str)
print(df["인원수"].dtypes)
df["인원수"] = df["인원수"].apply(lambda x : x + "명")
print(df)

# pivot_table 이용
data['value'] = 1
data['접수코드2'] = data['접수코드'].str.slice(start=0, stop=1)
data = pd.pivot_table(data, values='value', index=['접수코드2'], columns=['과목'], aggfunc=np.sum)
data.fillna(0, inplace=True)
print(data)

ws = pd.read_excel('C:/Users/kjy/Desktop/컴활/03-2010버전 출제유형분석3회/실습파일.xlsm', sheet_name=3, header=1)
ws['등록일자'] = ws['등록일자'].dt.month
ws['예상액'] = ws['수강료'] * 0.2
data = pd.pivot_table(ws, values=['수강료', '예상액'], index=['등록일자', '과목'], columns=['학년'], aggfunc=np.sum)
data.fillna('-', inplace=True)
