import pandas as pd
import statsmodels.api as sm


data1 = pd.read_excel('C:/Users/USER/Desktop/nightmare/新複製的每股盈餘變數.xlsx', header=[0,1], index_col=0)
print(data1.columns)
print(data1.head())
print(data1.iloc[0])
print("資料框的列標題層次：", data1.columns)
data1.columns = pd.MultiIndex.from_tuples([
    (pd.to_datetime(year), variables) for year, variables in data1.columns
])

all_years = data1.columns.get_level_values(0).year
target_years = range(2003, 2024)
target_variables = ['繼續營業單位損益', '資產總額', '普通股每股現金股利（盈餘及公積）', '總應計金額']

selected_data = data1.loc[:,
    all_years.isin(target_years) &
    data1.columns.get_level_values(1).isin(target_variables)
]

negearn = (selected_data.xs('繼續營業單位損益', level=1, axis=1) < 0).astype(int)
negearn.columns = pd.MultiIndex.from_product([negearn.columns, ['negearn']])  # 增加一層索引來命名 negearn

divdum = (selected_data.xs('普通股每股現金股利（盈餘及公積）', level=1, axis=1) > 0).astype(int)
divdum.columns = pd.MultiIndex.from_product([divdum.columns, ['divdum']])


result = pd.concat([selected_data, negearn, divdum], axis=1)

print(result.head())
print(result.columns)

future_predictions = pd.DataFrame(index=result.index, columns=result.columns)
future_predictions.loc[:, :] = result.loc[:, :]
print(future_predictions)

for t in range(2017, 2024):
    for tau in range(1, 6):  # 預測 t+1 到 t+5 年
        past_years = list(range(t - tau - 9, t - tau + 1))  # 回溯年份範圍 t-z-9 到 t-z
        historical_data = []

        print(f"Processing year: {t}, tau: {tau}, past_years: {past_years}")

        for i in past_years:
            try:
                X = result.xs(i, level=0, axis=1)[target_variables]
                X['divdum'] = divdum[i]
                X['negearn'] = negearn[i]
                y = result.xs(i + 1, level=0, axis=1)['繼續營業單位損益']  # 應變數
                X = sm.add_constant(X)
                model = sm.OLS(y, X).fit()
                historical_data.append(model.params)
            except KeyError as e:
                print(f"KeyError for year {i}: {e}")
                continue
            except Exception as e:
                print(f"Exception for year {i}: {e}")
                continue

        if historical_data:
            avg_coefficients = pd.DataFrame(historical_data).mean()
            X_future = result.xs(t, level=0, axis=1)[target_variables]
            X_future['divdum'] = divdum[t]
            X_future['negearn'] = negearn[t]
            X_future = sm.add_constant(X_future)
            future_predictions[f'Year_{t + tau}'] = X_future.dot(avg_coefficients)  # 預測未來盈餘
future_predictions.to_excel('C:/Users/USER/Desktop/future_earnings_predictions_adjusted.xlsx')
print("預測完成，結果已儲存於 'future_earnings_predictions_adjusted.xlsx'")
# 打印結果
#print(result.head())
#result.to_excel('C:/Users/USER/Desktop/result_negearn_divdum.xlsx')
# 將虛擬變數添加到 selected_data 中
#selected_data = pd.concat([selected_data, negearn.rename('negearn'), divdum.rename('divdum')], axis=1)

# 顯示篩選後的資料以及新虛擬變數
#print("篩選後的資料：")
#print(selected_data.head())
#selected_data.to_excel('C:/Users/USER/Desktop/selected_data.xlsx')


# 篩選MultiIndex中的年份和變數名稱
#selected_data = data1.loc[:,
#    data1.columns.get_level_values(0).isin([f"{year}-12-31" for year in target_years]) &
#    data1.columns.get_level_values('公司').isin(target_variables)
#]

#print("篩選後的資料：")
#print(selected_data.head())

#target1_years = list(range(2007-12-31, 2017-12-31))
#selected_data1 = data1.loc[:, data1.columns.isin(target1_years)]
#print(selected_data1)
# 再選取變數
#variables = ['繼續營業單位損益', '資產總額', '普通股每股現金股利(盈餘及公積)', '總應計金額']
#filtered_data = selected_data1.loc[:, selected_data1.columns.get_level_values('Variable').isin(variables)]
#target1_years = list(range(2007, 2017))
#selected_data1 = data1.loc[:, data1.columns.isin(target1_years)]
#print(selected_data1)

#variables = ['繼續營業單位損益', '資產總額', '普通股每股現金股利(盈餘及公積)', '總應計金額']

# 使用正則表達式篩選包含目標變數名稱的欄位
# 將變數名稱轉換為正則表達式的一部分，並進行篩選
#pattern = '|'.join(variables)
#filtered_data = selected_data1.filter(regex=pattern, axis=1)

# 顯示結果
#print("篩選後的資料：")
#print(filtered_data.head())



#variables = ['繼續營業單位損益', '資產總額', '普通股每股現金股利(盈餘及公積)', '總應計金額']

# 選取所需的資料
#selected_data1 = data1.loc[:, data1.columns.astype(str).isin(target_years)]
#selected_data1 = selected_data1[variables]
#print(selected_data1.head())





