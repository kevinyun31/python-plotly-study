#팬더 및 플롯 설치 - 아래 코드는 콘솔창에 입력한다
# pip install pandas
# pip install plotly

# --------------------------------------------------------------------------

#라이브러리를 가져오고 사용하기 쉽도록 별칭을 지정합니다.
import pandas as pd

import plotly.express as px

import plotly.io as pio

# --------------------------------------------------------------------------

#데이터세트를 읽고 변수 df에 저장
df = pd.read_excel('data.xlsx')

# 총 매출을 계산하고 이에 대한 새 열을 추가합니다.
df['Total Sales'] = df['Units Sold'] * df['Price per unit']

# --------------------------------------------------------------------------

# 제품 대 총 판매량의 막대 차트를 만듭니다.
fig = px.bar(df, x='Product', y='Total Sales', title='Product Sales')

# 차트 영역의 테두리와 배경색을 설정합니다.
fig.update_layout(
    plot_bgcolor='white',
    paper_bgcolor='lightgray',
    width=800,
    height=500,
    shapes=[dict(type='rect', xref='paper',
            yref='paper',
            x0=0,
            y0=0,
            x1=1,
            y1=1,
            line=dict(
                color='black',
                width=2,
            ),
        )
    ]
)

#그래프를 표시하다
fig.show()
print (f"첫번째")

pio.write_json(fig, 'figure.json', pretty=True)

# 아래 코드 줄을 사용하여 막대 그래프를 이미지에 저장
print (f"두번째")
pio.write_image(fig, 'bar2_graph.png')

print (f"세번째")

# --------------------------------------------------------------------------

# 엑셀 파일 생성 및 데이터 프레임 저장
# ExcelWriter는 내부에 writer.save()를 포함하고 있다
excel_path = 'report2.xlsx'
with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Sales Data')

    # 이미지를 엑셀에 추가
    worksheet = writer.sheets['Sales Data']
    worksheet.insert_image('H1', 'bar2_graph.png')

print(f"Excel 파일이 생성되었습니다: {excel_path}")

# --------------------------------------------------------------------------


