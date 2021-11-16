from xlrd import open_workbook

if __name__ == '__main__':

    # 读取4个工作簿（book）中的4个工作表（sheet）
    worksheets = [open_workbook('data/07.xls').sheet_by_index(0),
                  open_workbook('data/08.xls').sheet_by_index(0),
                  open_workbook('data/09.xls').sheet_by_index(0),
                  open_workbook('data/10.xls').sheet_by_index(0)]

    # 读入所有数据
    annual_situations = []
    for worksheet in worksheets:
        district = []
        category = []
        volume_of_trade = []
        for i in range(1, worksheet.nrows - 1):
            district.append(worksheet.cell_value(i, 2))
            category.append(worksheet.cell_value(i, 3))
            volume_of_trade.append(worksheet.cell_value(i, 7))
        annual_situation = [district, category, volume_of_trade]
        annual_situations.append(annual_situation)

    # 对 区县、类目 去重，得到 区县集合、类目集合
    all_district = []
    all_category = []
    for annual_situation in annual_situations:
        all_district.extend(annual_situation[0])
        all_category.extend(annual_situation[1])
    unique_district = list(set(all_district))
    unique_category = list(set(all_category))
    print(unique_district)
    print(unique_category)

