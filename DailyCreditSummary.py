# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import pandas as pd
import time
import datetime
import xlwings,json,requests,openpyxl
from io import BytesIO

gl_fw=45
off_days = {
    'Andy':           0,
    'Anna':           0,
    'Arthur':         0,
    "Bruno":          0,
    "Edwin":          0,
    "Hugh":           0,
    "Jack":           0,
    "Jeremy":         0,
    "Jerome":         0,
    "Jiaqi":          1,
    "Junsen":         0,
    "Kelly":          0,
    "Li":             0,
    "Mark":           1,
    "Niki":           1,
    "Nina":           0,
    "Qi":             1,
    "Qianqian":       0,
    "Sophia":         0,
    "Wan":            1,
    "Wuhao":          1,
    "Xuanyi":         1,
    "Lucas":          0
}
all_excel = r'C:\Users\lucashuang.FAREAST\OneDrive - Microsoft\Desktop\FY21_CaseAssignment.xlsx'
summary_excel = r'C:\Users\lucashuang.FAREAST\OneDrive - Microsoft\Desktop\summary.xlsx'


monitoring_fte = ['Andy', 'Anna', "Bruno","Sophia","Hugh",
                 "Junsen", "Kelly", "Mark",
                 "Niki", "Nina", "Qi", "Qianqian", "Wuhao","Howard","Jimmy","Lucas"]
monitoring_vendor = ['Arthur',  "Edwin", "Aristo",
                 "Jack", "Jerome", "Jiaqi",  "Li",
                 "Wan"]

name_mapping={"Nina Li":"Nina","Maggie Dong":"Maggie","Anna Gao":"Anna","Andy Wu":"Andy","Yinan Zhou":"Kelly","Qi Chen":"Qi",
              "Wuhao Chen":"Wuhao","Qianqian Liu":"Qianqian","Junsen Chen":"Junsen","Mark He":"Mark","Hugh Chao":"Hugh",
               "Sophia Zhang":"Sophia","Hao Pei":"Howard","Ji Bian":"Jimmy","Niki Jin":"Niki","Wan Huang":"Wan","Li Zhang":"Li",
              "Yue Mei":"Edwin","Jeremy Liang":"Jeremy","Jack Bian":"Jack","Jiaqi Deng":"Jiaqi","Arthur Huang":"Arthur",
              "Junhao Guan":"Jerome","Lucas Huang":"Lucas","Xuanyi Liu":"Xuanyi","Chener Zhang":"Chener","Fang Liao":"Aristo"}

pd.set_option('display.max_columns', None)
# ???????????????
pd.set_option('display.max_rows', None)


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.

def get_week_off():
    url = "https://lucasfunctionappanalyzeresourceschedule2.azurewebsites.net/api/LucasFunctionAppAnalyzeResourceScheduleHTTPTrigger"
    payload = json.dumps({"date_list":["123"],"command":"week_until_today"})
    response = requests.post(url, data=payload)
    week_off_dict = json.loads(response.text)
    new_week_off_dict = {}
    for k,v in week_off_dict.items():
        if k in name_mapping.keys():
            name = name_mapping[k]
            if name in monitoring_fte or name in monitoring_vendor:
                new_week_off_dict[name]=v
    return new_week_off_dict
# Press the green button in the gutter to run the script.
def get_excel_data():

    method =0
    if method:
        # ??????????????????????????????????????????
        df = pd.read_excel(all_excel,
                           sheet_name='Azure Monitoring',engine='openpyxl')
        # ?????????????????????
        df = df.dropna(axis=1, how='all')
        df = df.loc[df['FW'] == gl_fw]
        return df
    else:
        # ??????????????????????????????????????????
        response = requests.get('https://prod-00.eastus.logic.azure.com:443/workflows/764229174611433581f584080a1c15c1/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=CekobRDm-H9Dx-tBTpXOblRXrJqihxBoTeCyilDCi2w')
        excel_response = requests.get("https://lucasstorageaccount.blob.core.windows.net/token/Monitoring Today's Cases and Credit This Week.xlsx")
        excel_response = excel_response.content
        with open("./Monitoring Today's Cases and Credit This Week.xlsx",'wb') as f:
            f.write(excel_response)
        df = pd.read_excel(r"./Monitoring Today's Cases and Credit This Week.xlsx",
                           sheet_name='Case this week',engine='openpyxl')
        # ?????????????????????
        df = df.dropna(axis=0, how='all')
        #df = df.dropna(axis=1, how='all')
        return df


def get_se_data(df_excel,monitoring_se):
    d = time.localtime()

    weekday_shift = time.strftime("%w", d)
    weekday_shift = int(weekday_shift)
    today = time.strftime("%Y-%m-%d", d)
    print("today:",today)
    print("weekday_shift = ",weekday_shift)

    ret = pd.DataFrame(columns=['workdays', 'case today', 'task today',
                                "active case this week", "active task this week",
                                "active credit per day", "transferred out case", "transferred out task"],
                       index=monitoring_se)
    ret.loc[:, :] = 0
    ret.loc[:, "workdays"] = weekday_shift
    for se in monitoring_se:
        # workdays
        #ret.loc[se, "workdays"] -= off_days[se]
        if se in new_week_off_dict.keys():
            ret.loc[se, "workdays"] -= new_week_off_dict[se]
        # case this week
        temp_df_case_this_week = df_excel[df_excel['Case/Task'].str.contains("Case")]
        temp_df_case_this_week = temp_df_case_this_week[temp_df_case_this_week["Case owner"] == se]
        ret.loc[se, "active case this week"] = temp_df_case_this_week.shape[0]
        # task this week
        temp_df_task_this_week = df_excel[df_excel['Case/Task'].str.contains("Task")]
        temp_df_task_this_week = temp_df_task_this_week[temp_df_task_this_week["Case owner"] == se]
        ret.loc[se, "active task this week"] = temp_df_task_this_week.shape[0]
        # case today
        # temp_df_case_today = df_excel[df_excel['Case/Task'].str.contains("Case")]
        # temp_df_case_today = temp_df_case_today[temp_df_case_today["Weekday Shift"] == weekday_shift]
        # temp_df_case_today = temp_df_case_today[temp_df_case_today["Case owner"] == se]
        temp_df_case_today = df_excel[df_excel['Case/Task'].str.contains("Case")]
        temp_df_case_today = temp_df_case_today[temp_df_case_today["Date"] == today]
        temp_df_case_today = temp_df_case_today[temp_df_case_today["Case owner"] == se]
        ret.loc[se, "case today"] = temp_df_case_today.shape[0]
        # task today
        # temp_df_task_today = df_excel[df_excel['Case/Task'].str.contains("Task")]
        # temp_df_task_today = temp_df_task_today[temp_df_task_today["Weekday Shift"] == weekday_shift]
        # temp_df_task_today = temp_df_task_today[temp_df_task_today["Case owner"] == se]
        temp_df_task_today = df_excel[df_excel['Case/Task'].str.contains("Task")]
        temp_df_task_today = temp_df_task_today[temp_df_task_today["Date"] == today]
        temp_df_task_today = temp_df_task_today[temp_df_task_today["Case owner"] == se]
        ret.loc[se, "task today"] = temp_df_task_today.shape[0]
        # transfer out case
        temp_df_transfer_out_case = df_excel[df_excel['Case/Task'].str.contains("Case")]
        temp_df_transfer_out_case = temp_df_transfer_out_case[temp_df_transfer_out_case["Case owner"] == se]
        temp_df_transfer_out_case = temp_df_transfer_out_case[temp_df_transfer_out_case["Transfer Out?"] == "Y"]
        ret.loc[se, "transferred out case"] = temp_df_transfer_out_case.shape[0]
        ##transfer out task
        temp_df_transfer_out_task = df_excel[df_excel['Case/Task'].str.contains("Task")]
        temp_df_transfer_out_task = temp_df_transfer_out_task[temp_df_transfer_out_task["Case owner"] == se]
        temp_df_transfer_out_task = temp_df_transfer_out_task[temp_df_transfer_out_task["Transfer Out?"] == "Y"]
        ret.loc[se, "transferred out task"] = temp_df_transfer_out_task.shape[0]
        # minus transfer out
        ret.loc[se, "active case this week"] -= ret.loc[se, "transferred out case"]
        ret.loc[se, "active task this week"] -= ret.loc[se, "transferred out task"]
        # credit
        if ret.loc[se, "workdays"] == 0:
            ret.loc[se, "active credit per day"] = 0
        else:
            ret.loc[se, "active credit per day"] = (ret.loc[se, "active case this week"] +
                                                    ret.loc[se, "active task this week"] * 0.5) / ret.loc[
                                                       se, "workdays"]
    ret.sort_values("case today", ascending=False, inplace=True)
    ret.sort_values("workdays", ascending=False, inplace=True)
    ret.sort_values("active credit per day", ascending=False, inplace=True)
    return ret





def output_excel(df_all):
    print("===output excel====")
    today = str(datetime.date.today().month) + "." + str(datetime.date.today().day)
    print("today =",today)
    wb = xlwings.Book(summary_excel)
    try:
        sht = wb.sheets.add(name=today, before=None, after=None)
    except:
        sht = wb.sheets[today]
    sht.range("A1").value = df_all
    wb.save()


def concat_and_sort(df_fte, df_vendor):
    df_fte=pd.concat([df_fte,df_vendor])
    sum_case_today = df_fte["case today"].sum()
    sum_task_today = df_fte["task today"].sum()
    sum_active_case_this_week = df_fte["active case this week"].sum()
    sum_active_task_this_week = df_fte["active task this week"].sum()
    df_fte.loc["Total"] = ""
    df_fte.loc["Total", "case today"] = sum_case_today
    df_fte.loc["Total", "task today"] = sum_task_today
    df_fte.loc["Total", "active case this week"] = sum_active_case_this_week
    df_fte.loc["Total", "active task this week"] = sum_active_task_this_week
    return df_fte

if __name__ == '__main__':
    print_hi('PyCharm')
    new_week_off_dict = get_week_off()
    df_excel = get_excel_data()
    print(df_excel)
    df_fte = get_se_data(df_excel,monitoring_fte)
    print(df_fte)
    df_vendor = get_se_data(df_excel,monitoring_vendor)
    print(df_vendor)
    df_all = concat_and_sort(df_fte,df_vendor)
    output_excel(df_all)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
