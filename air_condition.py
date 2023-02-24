import xlwings as xw
from xlwings.utils import rgb_to_int
from xlwings.utils import rgb_to_hex
import pandas as pd
import json
import requests
import numpy as np
import matplotlib.pyplot as plt

def get_data(sheet):

    url = 'https://data.epa.gov.tw/api/v2/aqx_p_432?api_key=e8dd42e6-9b8b-43f8-991e-b3dee723a52d&limit=1000&sort=ImportDate%20desc&format=JSON'
    response = requests.get(url)
    data = json.loads(response.text)
    df = pd.DataFrame(data['records'])
    col_dict = {}
    for f in data['fields']:
        col_dict[f['id']] = f['info']['label']
        col_dict['publishtime'] = '更新時間'

    df = df.rename(columns=col_dict)
    df['空氣品質指標'].replace('', np.nan, inplace=True)
    df.dropna(subset=['空氣品質指標'], inplace=True)
    sheet['A1'].options(index=False).value = df

def set_style(sheet):
    sheet.autofit()
    sheet['A:A'].insert()
    sheet['A1'].options(transpose=True).value = sheet['C1'].expand('down').value
    sheet['C1'].api.EntireColumn.Delete()
    for cell in sheet['C2'].expand('down'):

        cell.api.Font.Color = rgb_to_int((255, 255, 255))
        #     cell.api.font_object.color.set(rgb_to_int((255, 255, 255)))

        if cell.value <= 50:
            cell.color = (0, 162, 63)  # 良好(綠色)

        elif cell.value <= 100:
            cell.color = (253, 217, 1)  # 普通(黃色)
            cell.api.Font.Color = rgb_to_int((0, 0, 0))
            # cell.api.font_object.color.set(rgb_to_int((0, 0, 0)))
        elif cell.value <= 150:
            cell.color = (228, 120, 9)  # 對敏感群族不良(橘色)
        elif cell.value <= 200:
            cell.color = (228, 0, 21)  # 對所有群族不健康(紅色)
        elif cell.value <= 300:
            cell.color = (174, 3, 133)  # 非常不健康(紫色)
        else:
            cell.color = (165, 30, 52)  # 危害(棕色)

def statistic(sheet,sheet2):

    df = sheet['A1'].options(convert=pd.DataFrame, expand='table', index=False).value
    main_df = pd.DataFrame()
    main_df['監測站數'] = df.groupby('縣市').size()
    main_df['最大值'] = df.groupby('縣市')['空氣品質指標'].max()
    main_df['最小值'] = df.groupby('縣市')['空氣品質指標'].min()
    main_df['平均值'] = df.groupby('縣市')['空氣品質指標'].mean().round(1)
    df.groupby('縣市')['狀態'].apply(list)
    category = ['良好', '普通', '對敏感族群不健康']
    for i in category:
        main_df[i] = df.groupby('縣市')['狀態'].apply(lambda x: (x == i).sum())

    main_df = main_df.sort_values(by='平均值', ascending=True)
    sheet2['B2'].value = main_df

    for cell in sheet2['F3'].expand('down'):

        cell.api.Font.Color = rgb_to_int((255, 255, 255))
    #     cell.api.font_object.color.set(rgb_to_int((255, 255, 255)))

        if cell.value <= 50:
            cell.color = (0, 162, 63)  # 良好(綠色)

        elif cell.value <= 100:
            cell.color = (253, 217, 1)  # 普通(黃色)
            cell.api.Font.Color = rgb_to_int((0, 0, 0))
        # cell.api.font_object.color.set(rgb_to_int((0, 0, 0)))
        elif cell.value <= 150:
            cell.color = (228, 120, 9)  # 對敏感群族不良(橘色)
        elif cell.value <= 200:
            cell.color = (228, 0, 21)  # 對所有群族不健康(紅色)
        elif cell.value <= 300:
            cell.color = (174, 3, 133)  # 非常不健康(紫色)
        else:
            cell.color = (165, 30, 52)  # 危害(棕色)
    sheet2['N2'].value = df['空氣品質指標'].describe()
    sheet2.autofit()
def plot_fig(sheet):
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
    df = sheet['B2'].options(convert=pd.DataFrame, expand='table').value
    bar_color =[]

    for i in df['平均值']:
        if i <= 50:
            bar_color.append(rgb_to_hex(0, 162, 63))  # 良好(綠色)
        elif i <= 100:
            bar_color.append(rgb_to_hex(253, 217, 1))  # 普通(黃色)
        elif i <= 150:
            bar_color.append(rgb_to_hex(228, 120, 9))  # 對敏感群族不良(橘色)

    ax = df['平均值'].plot(kind='bar', color= bar_color, figsize=(7, 4))
    plt.title('空氣品質', fontsize=14)
    plt.xlabel('平均值', fontsize=12)
    plt.ylabel('縣市', fontsize=12)
    fig = ax.get_figure()
    plot = sheet.pictures.add(fig, name='fig', update=True)
    plot.left = sheet['Q2'].left
    plot.top = sheet['Q2'].top

def run():
    wb = xw.Book()
    sheet = wb.sheets[0]
    sheet.name = '空氣品質清單'
    get_data(sheet)
    set_style(sheet)
    sheet2 = wb.sheets.add('統計資料', after='空氣品質清單')
    statistic(sheet, sheet2)
    plot_fig(sheet2)
    wb.save()

run()
