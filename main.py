import numpy as np
from xlrd import open_workbook
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

if __name__ == '__main__':

    # 读取4个工作簿（book）中的4个工作表（sheet）
    worksheets = [open_workbook('data/07.xls').sheet_by_index(0),
                  open_workbook('data/08.xls').sheet_by_index(0),
                  open_workbook('data/09.xls').sheet_by_index(0),
                  open_workbook('data/10.xls').sheet_by_index(0)]

    # 读入所有数据
    monthly_situations = []
    for worksheet in worksheets:
        district = []
        category = []
        volume_of_trade = []
        for i in range(1, worksheet.nrows - 1):
            district.append(worksheet.cell_value(i, 2))
            category.append(worksheet.cell_value(i, 3))
            volume_of_trade.append(float(worksheet.cell_value(i, 8)))
        monthly_situation = [district, category, volume_of_trade]
        monthly_situations.append(monthly_situation)

    # 对 区县、类目 去重，得到 区县集合、类目集合
    all_district = []
    all_category = []
    for monthly_situation in monthly_situations:
        all_district.extend(monthly_situation[0])
        all_category.extend(monthly_situation[1])
    unique_district = list(set(all_district))
    unique_category = list(set(all_category))
    print(unique_district)
    print(unique_category)

    m_c_v = []
    for monthly_situation in monthly_situations:
        c_v = [0 for i in range(len(unique_category))]
        category = monthly_situation[1]
        volume_of_trade = monthly_situation[2]
        for i in range(len(category)):
            index = unique_category.index(category[i])
            c_v[index] += volume_of_trade[i]
        m_c_v.append(c_v)
    print(m_c_v)

    # 构造需要显示的值
    X = [7, 8, 9, 10]  # X轴的坐标
    Y = range(len(unique_category))  # Y轴的坐标
    # 设置每一个（X，Y）坐标所对应的Z轴的值，在这边Z（X，Y）=X+Y
    Z = np.zeros(shape=(4, len(unique_category)))
    for i in range(4):
        for j in range(len(unique_category)):
            Z[i, j] = i + j
    print(Z)

    for i in range(4):
        for j in range(len(unique_category)):
            Z[i, j] = m_c_v[i][j]

    xx, yy = np.meshgrid(X, Y)  # 网格化坐标
    X, Y = xx.ravel(), yy.ravel()  # 矩阵扁平化
    bottom = np.zeros_like(X)  # 设置柱状图的底端位值
    Z = Z.ravel()  # 扁平化矩阵

    width = height = 1  # 每一个柱子的长和宽

    # 绘图设置
    fig = plt.figure(figsize=(6.4, 4.8), dpi=100)
    ax = fig.gca(projection='3d')  # 三维坐标轴
    ax.get_proj = lambda: np.dot(Axes3D.get_proj(ax), np.diag([1, 1.5, 1, 1]))
    ax.bar3d(X, Y, bottom, width, height, Z, shade=True)  #
    # 坐标轴设置
    ax.set_xlabel('Month')
    ax.set_ylabel('index of SPU')
    ax.set_zlabel('Volume of trade')
    plt.show()
    print('ok')
