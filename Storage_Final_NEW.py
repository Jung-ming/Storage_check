import datetime
import pandas as pd
import Images
import csv


# '客戶\n名稱\nSource'
# print(庫存報表.columns)
def 檔案核對(銷貨報表, 庫存報表, 月結客戶, 當天日期):
    銷貨報表 = pd.read_excel(f'{銷貨報表}', header=0)  # '料件編號'
    庫存報表 = pd.read_excel(f'{庫存報表}')  # '研騰件號\nPSI Part No.'
    輸出用檔案 = pd.DataFrame()
    輸出用檔案 = 輸出用檔案.assign(料件編號=None, 客戶=None, 母工單單號=None, MO=None, 銷貨紀錄=None, GDS=None)

    足標記數 = 0
    for index1, value1 in 銷貨報表.iterrows():
        # 先檢查銷貨報表的料件編號是否在裡面，如果不在裡面，就填入編號但不進行後續的判斷
        if value1['料件編號'] not in list(庫存報表['研騰件號\nPSI Part No.']):
            輸出用檔案.at[足標記數, '料件編號'] = value1['料件編號']
            輸出用檔案.at[足標記數, 'GDS'] = '已出貨'
            足標記數 += 1
        else:
            待寫入資料 = 庫存報表[庫存報表['研騰件號\nPSI Part No.'] == value1['料件編號']]
            for index2, value2 in 待寫入資料.iterrows():
                # 若 'MIS Ship Remark' 銷貨紀錄 和 'GDS' 庫存皆為空
                # 則跳過此資料
                if pd.isna(value2['MIS Ship Remark']) and pd.isna(value2['GDS']):
                    continue
                else:
                    輸出用檔案.at[足標記數, '客戶'] = value2['客戶\n名稱\nSource']
                    輸出用檔案.at[足標記數, '料件編號'] = value1['料件編號']
                    輸出用檔案.at[足標記數, '母工單單號'] = value2['母工單單號']
                    輸出用檔案.at[足標記數, 'MO'] = value2['客戶MO. No \nCustomer MO.']
                    輸出用檔案.at[足標記數, '銷貨紀錄'] = value2['MIS Ship Remark']
                    輸出用檔案.at[足標記數, 'GDS'] = value2['GDS']
                    足標記數 += 1

    # for 足標, 欄位 in 輸出用檔案.iterrows():
    #     # 非月結客戶只保留核對當天的出貨紀錄
    #     # 除非GDS有庫存，否則沒有當天的出貨紀錄就表示這筆資料通常不需要
    #     if 欄位['客戶'] not in 月結客戶:
    #         if pd.isna(欄位['GDS']) and 當天日期 not in str(欄位['銷貨紀錄']):
    #             輸出用檔案 = 輸出用檔案.drop(足標)

    for 足標, 欄位 in 銷貨報表.iterrows():
        if 欄位['料件編號'] not in list(輸出用檔案['料件編號']):
            輸出用檔案.at[足標記數, '料件編號'] = 欄位['料件編號']
            足標記數 += 1

    return 輸出用檔案


def Excel格式輸出(輸出用檔案):
    writer = pd.ExcelWriter('銷貨核對-New.xlsx', engine='xlsxwriter')
    輸出用檔案.to_excel(writer, index=False, sheet_name='銷貨核對')

    worksheet = writer.sheets['銷貨核對']

    # 格式設定
    銷貨紀錄格式 = writer.book.add_format({'font_size': 11, 'text_wrap': True})
    一般欄位格式 = writer.book.add_format({'font_size': 11, 'valign': 'vcenter', 'text_wrap': True})
    # 欄寬設置
    # 料件編號 18
    # 銷貨紀錄 75
    # 母工單單號 13
    worksheet.set_column(0, 0, 18, 一般欄位格式)
    worksheet.set_column(1, 1, 13, 一般欄位格式)
    worksheet.set_column(2, 2, 13, 一般欄位格式)
    worksheet.set_column(3, 3, 15, 一般欄位格式)
    worksheet.set_column(4, 4, 75, 銷貨紀錄格式)

    # 使用自訂的標題格式
    for 欄位, 欄位名 in enumerate(list(輸出用檔案.columns)):
        worksheet.write(0, 欄位, 欄位名, 一般欄位格式)
    writer.save()


if __name__ == "__main__":
    當天日期_日期格式 = datetime.date.today()
    當天日期_文字格式 = 當天日期_日期格式.strftime('%y/%m/%d')

    月結客戶 = ['2E-營邦', '7O-兆赫', '44-Asentria', '5Q-B&B', '0K-KFA']

    輸出用檔案 = 檔案核對('GDS.xls', 'csfp422-241011-151059.xls', 月結客戶, 當天日期_文字格式)
    Excel格式輸出(輸出用檔案)
