# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import pandas as pd
import xlwings
import requests

#################Parameters######################
gl_fw=33
backlog_excel = r'D:\DailyCopy\backlog-fw'+str(gl_fw)+'.xlsx'
#case_excel = r'C:\Users\v-zixinh\OneDrive - Microsoft\Desktop\FY21_CaseAssignment.xlsx'
generate_excel = r'D:\DailyCopy\output-fw'+str(gl_fw)+'.xlsx'
#################################################

monitoring_vendor_se = ['Arthur', "Edwin",
                 "Jack", "Jerome", "Jiaqi", "Li", "Wan", "Aristo"]

monitoring_fte_se = ['Andy', 'Anna', "Bruno",  "Junsen", "Kelly",  "Mark",
                 "Niki", "Nina", "Qi", "Qianqian",  "Wuhao","Hugh","Sophia","Howard","Jimmy","Lucas"]

integration_se = ["yuzhang6","yuaf","jiecao","zhangz","qili7","v-xuanyiliu","yinduoli","yanden","yinshi","huidongliu","beixiao"
                  ]

#all_se = monitoring_fte_se + monitoring_vendor_se
all_se = monitoring_fte_se + monitoring_vendor_se + integration_se

possible_names = {"Andy":["Andy Wu","Hao Wu","Andy W"],"Anna":["Xue Gao","Xue G"],"Bruno":["Bruno L"],"Hugh":["Hui C"],"Junsen":["Junsen C"],
                 "Kelly":["Yinan Zhou","Yinan Z"],"Qianqian":["Qianqian l","Qianqian liu"],"Maggie":["Meijiao Dong","Maggie D"],
                 "Mark":["Xiaowei He","Xiaowei H"],"Nina":["Na L"],"Qi":["Qi C"],"Sophia":["Sophia Z"],"Arthur":["Arthur Huang","Arthur H"],
                 "Jack":["Jack Bian"],"Jeremy":["Jeremy Liang"],"Jerome":["Junhao Guan","Junhao G","Jerome G","Jerome Guan"],"Jiaqi":["Jiaqi Deng"],"Li":["Li Zhang"],
                 "Wan":["Treasure Huang","Treasure H"],"Edwin":["Edwin Mei","Edwin M"],"Lucas":["Zixin H"],"Wuhao":["Wuhao Chen","Wuhao C"],
                 "Niki":["Yan J"],"Howard":["Howard P"],"Jimmy":["Ji B"],"Chener":["Chener Zhang","Chener Z"],
                  "Aristo":["Fang L","Fang Liao"],
                  "yuzhang6":["Yu Zhang","Yu Z"],"v-xuanyiliu":["Xuanyi L","v-xuanyiliu"],"jiecao":["Jie Cao","Jie C"],"qili7":["Qing L"],
                  "yanden":["Yanbo Deng","Yanbo D"],"yinshi":["Yingjie Shi","Yingjie S"],
                  "yuaf":["Yuanchang F"],"zhangz":["Ziyu Z"],"huidongliu":["Huidong Liu","Huidong L"],"beixiao":["Bei Xiao","Bei X"],"yinduoli":["YINDUO L","Yinduo Li"]}

# 显示所有行
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
def get_excel_data(fw):



    # 读取工作簿和工作簿中的工作表
    df_backlog = pd.read_excel(backlog_excel,engine='openpyxl')
    df_backlog = df_backlog.dropna(axis=1, how='all')
    df_backlog = df_backlog[["Names","Cases","All Items"]]
    print("------------backlog=-------------")
    print(df_backlog)
    # 获取在线case assignemnt
    response = requests.get("https://prod-05.southeastasia.logic.azure.com:443/workflows/30225d47c0024af3a79b0b2f2c4ab996/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=J3vDhtu19ZbZdd5J_lggNfAsVd5PSncOkpFxESINwrg")
    excel_response = requests.get(
        "https://caseassignmentblob.blob.core.windows.net/casevalumeblob1/CaseAssignment.xlsx")
    excel_response = excel_response.content
    with open("./CaseAssignment.xlsx", 'wb') as f:
        f.write(excel_response)
    case_excel = "./CaseAssignment.xlsx"
    df_monitoring_case = pd.read_excel(case_excel,sheet_name='Azure Monitoring',engine='openpyxl')
    df_monitoring_case = df_monitoring_case.dropna(axis=1, how='all')
    df_monitoring_case = df_monitoring_case.loc[df_monitoring_case['FW'] == fw]
    print("------------df_monitoring_case=-------------")

    print(df_monitoring_case)
    df_integration_case = pd.read_excel(case_excel,sheet_name='Integration',engine='openpyxl')
    df_integration_case = df_integration_case.dropna(axis=1, how='all')
    df_integration_case = df_integration_case.loc[df_integration_case['FW'] == fw]
    print("------------df_integration_case=-------------")

    print(df_integration_case)
    # df_case = pd.concat([df_monitoring_case,df_integration_case])
    # df_case = df_case.dropna(axis=1, how='all')
    #
    # print(df_case)
    # 新建一个工作簿
    ret = pd.DataFrame(columns=['Case Volumn', 'Collaboration Task Volume', 'Escalation Task Volume',
                                "Follow-up Task Volume", "Rave AR",
                                "Weekly Total", "Engineers Backlog (No Rave)", "Engineers Backlog Case Volume",
                                "Engineers Backlog Task Volume"],
                       index=all_se)
    ret.loc[:, :] = 0

    ret_general = pd.DataFrame(columns=["PoD Name","Case Volume","CritSit Case Volume","ARR Case Volume",
                                        "Collaboration Task Volume","Escalation Task Volume","Follow-up Task Volume",
                                        "Rave AR","Weekly Total","Engineers Backlog (No Rave)","Engineers Backlog Case Volume",
                                        "Engineers Backlog Task Volume"])
    ret_general.loc[:, :] = 0
    ret_general.loc[0,"PoD Name"]="Monitoring"
    ret_general.loc[1,"PoD Name"]="Integration"
    #general crisit case volumn
    temp_df_monitoring_crisit_case_this_week = df_monitoring_case[df_monitoring_case['Case/Task'].str.lower().str.contains("case")]
    temp_df_monitoring_crisit_case_this_week = temp_df_monitoring_crisit_case_this_week[temp_df_monitoring_crisit_case_this_week['Severity'].str.lower().str.contains("a")]
    ret_general.loc[0, "CritSit Case Volume"] = temp_df_monitoring_crisit_case_this_week.shape[0]

    temp_df_integration_crisit_case_this_week = df_integration_case[df_integration_case['Case/Task'].str.lower().str.contains("case")]
    temp_df_integration_crisit_case_this_week = temp_df_integration_crisit_case_this_week[temp_df_integration_crisit_case_this_week['Severity'].str.lower().str.contains("a")]
    ret_general.loc[1, "CritSit Case Volume"] = temp_df_integration_crisit_case_this_week.shape[0]
    # general ARR case volumn
    temp_df_monitoring_arr_case_this_week = df_monitoring_case[df_monitoring_case['Case/Task'].str.lower().str.contains("case")]
    temp_df_monitoring_arr_case_this_week = temp_df_monitoring_arr_case_this_week[temp_df_monitoring_arr_case_this_week['ARR/Unified/Premier/Pro?'].str.lower().str.contains("arr")]
    ret_general.loc[0, "ARR Case Volume"] = temp_df_monitoring_arr_case_this_week.shape[0]

    temp_df_integration_arr_case_this_week = df_integration_case[df_integration_case['Case/Task'].str.lower().str.contains("case")]
    temp_df_integration_arr_case_this_week = temp_df_integration_arr_case_this_week[temp_df_integration_arr_case_this_week['ARR/Unified/Premier/Pro?'].str.lower().str.contains("arr")]
    ret_general.loc[1, "ARR Case Volume"] = temp_df_integration_arr_case_this_week.shape[0]

    # general case volumn
    temp_df_monitoring_case_this_week = df_monitoring_case[
        df_monitoring_case['Case/Task'].str.lower().str.contains("case")]
    ret_general.loc[0, "Case Volume"] = temp_df_monitoring_case_this_week.shape[0]

    temp_df_integration_case_this_week = df_integration_case[
        df_integration_case['Case/Task'].str.lower().str.contains("case")]
    ret_general.loc[1, "Case Volume"] = temp_df_integration_case_this_week.shape[0]

    # general collab volumn
    temp_df_monitoring_collab_this_week = df_monitoring_case[
        df_monitoring_case['Case/Task'].str.lower().str.contains("collab")]
    ret_general.loc[0, "Collaboration Task Volume"] = temp_df_monitoring_collab_this_week.shape[0]

    temp_df_integration_collab_this_week = df_integration_case[
        df_integration_case['Case/Task'].str.lower().str.contains("collab")]
    ret_general.loc[1, "Collaboration Task Volume"] = temp_df_integration_collab_this_week.shape[0]

    # general follow up volumn
    temp_df_monitoring_follow_this_week = df_monitoring_case[
        df_monitoring_case['Case/Task'].str.lower().str.contains("follow")]
    ret_general.loc[0, "Follow-up Task Volume"] = temp_df_monitoring_follow_this_week.shape[0]

    temp_df_integration_follow_this_week = df_integration_case[
        df_integration_case['Case/Task'].str.lower().str.contains("follow")]
    ret_general.loc[1, "Follow-up Task Volume"] = temp_df_integration_follow_this_week.shape[0]
    # general rave volumn
    temp_df_monitoring_rave_this_week = df_monitoring_case[
        df_monitoring_case['Case/Task'].str.lower().str.contains("rave")]
    ret_general.loc[0, "Rave AR"] = temp_df_monitoring_rave_this_week.shape[0]

    temp_df_integration_rave_this_week = df_integration_case[
        df_integration_case['Case/Task'].str.lower().str.contains("rave")]
    ret_general.loc[1, "Rave AR"] = temp_df_integration_rave_this_week.shape[0]

    # general weekly total volumn
    ret_general.loc[0, "Weekly Total"] = ret_general.loc[0, "Case Volume"] + ret_general.loc[0, "Collaboration Task Volume"] \
                                        +ret_general.loc[0, "Follow-up Task Volume"]+ret_general.loc[0, "Rave AR"]

    ret_general.loc[1, "Weekly Total"] = ret_general.loc[1, "Case Volume"] + ret_general.loc[1, "Collaboration Task Volume"] \
                                         + ret_general.loc[1, "Follow-up Task Volume"] + ret_general.loc[1, "Rave AR"]

    for se in all_se:
        print(" ------------- ",se,'-----------------')
        if se in integration_se:
            df_case = df_integration_case
        else:
            df_case = df_monitoring_case

        # case this week
        temp_df_case_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("case")]
        temp_df_case_this_week = temp_df_case_this_week[temp_df_case_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Case Volumn"] = temp_df_case_this_week.shape[0]
        # collab this week
        temp_df_task_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("collab")]
        temp_df_task_this_week = temp_df_task_this_week[temp_df_task_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Collaboration Task Volume"] = temp_df_task_this_week.shape[0]

        # Follow-up this week
        temp_df_followup_task_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("follow")]
        temp_df_followup_task_this_week = temp_df_followup_task_this_week[temp_df_followup_task_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Follow-up Task Volume"] = temp_df_followup_task_this_week.shape[0]
        # Rave this week
        temp_df_rave_task_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("rave")]
        temp_df_rave_task_this_week = temp_df_rave_task_this_week[temp_df_rave_task_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Rave AR"] = temp_df_rave_task_this_week.shape[0]
        # weekly total
        ret.loc[se, "Weekly Total"] = ret.loc[se, "Case Volumn"]+ret.loc[se, "Follow-up Task Volume"] \
                                    +ret.loc[se, "Rave AR"] + ret.loc[se, "Collaboration Task Volume"]
        se_name=''
        for name in possible_names[se]:
            if name in df_backlog.Names.values:
                se_name=name
                print(" ------------- ",se_name,'-----------------')
                # Engineers Backlog Case Volume
                temp_df_backlog_case = df_backlog[df_backlog["Names"].str.strip() == se_name]
                print(temp_df_backlog_case)
                try:
                    ret.loc[se, "Engineers Backlog Case Volume"] += temp_df_backlog_case["Cases"].values[0]
                except IndexError:
                    ret.loc[se, "Engineers Backlog Case Volume"] = 0

                # Engineers Backlog (No Rave)
                temp_df_backlog_all = df_backlog[df_backlog["Names"].str.strip() == se_name]
                try:
                    ret.loc[se, "Engineers Backlog (No Rave)"] += temp_df_backlog_all["All Items"].values[0]
                except IndexError:
                    ret.loc[se, "Engineers Backlog (No Rave)"] = 0
        #Engineers Backlog Task Volume
        ret.loc[se, "Engineers Backlog Task Volume"] =  ret.loc[se, "Engineers Backlog (No Rave)"] - ret.loc[se, "Engineers Backlog Case Volume"]

    ret_monitoring_vendor = ret.loc[monitoring_vendor_se]
    ret_monitoring_vendor.sort_values("Engineers Backlog (No Rave)", ascending=False, inplace=True)
    ret_monitoring_vendor.sort_values("Weekly Total", ascending=False, inplace=True)

    ret_monitoring_fte = ret.loc[monitoring_fte_se]
    ret_monitoring_fte.sort_values("Engineers Backlog (No Rave)", ascending=False, inplace=True)
    ret_monitoring_fte.sort_values("Weekly Total", ascending=False, inplace=True)

    ret_integration = ret.loc[integration_se]
    ret_integration.sort_values("Engineers Backlog (No Rave)", ascending=False, inplace=True)
    ret_integration.sort_values("Weekly Total", ascending=False, inplace=True)

    # general all backlog volumn
    ret_general.loc[0, "Engineers Backlog (No Rave)"] = ret_monitoring_fte["Engineers Backlog (No Rave)"].sum()+ret_monitoring_vendor["Engineers Backlog (No Rave)"].sum()
    ret_general.loc[1, "Engineers Backlog (No Rave)"] = ret_integration["Engineers Backlog (No Rave)"].sum()

    # general backlog case volumn
    ret_general.loc[0, "Engineers Backlog Case Volume"] = ret_monitoring_fte["Engineers Backlog Case Volume"].sum() + \
                                                        ret_monitoring_vendor["Engineers Backlog Case Volume"].sum()
    ret_general.loc[1, "Engineers Backlog Case Volume"] = ret_integration["Engineers Backlog Case Volume"].sum()
    # general backlog task volumn
    ret_general.loc[0, "Engineers Backlog Task Volume"] = ret_monitoring_fte["Engineers Backlog Task Volume"].sum() + \
                                                        ret_monitoring_vendor["Engineers Backlog Task Volume"].sum()
    ret_general.loc[1, "Engineers Backlog Task Volume"] = ret_integration["Engineers Backlog Task Volume"].sum()

    print('path = ',generate_excel)
    wb = xlwings.Book(generate_excel)
    try:
        sht = wb.sheets.add(name="Monitoring FTE", before=None, after=None)
    except:
        sht = wb.sheets["Monitoring FTE"]
        sht.clear()

    sht.range("A1").value = ret_monitoring_fte

    try:
        sht = wb.sheets.add(name="Monitoring Vendor", before=None, after=None)
    except:
        sht = wb.sheets["Monitoring Vendor"]
        sht.clear()
    sht.range("A1").value = ret_monitoring_vendor
    try:
        sht = wb.sheets.add(name="Integration", before=None, after=None)
    except:
        sht = wb.sheets["Integration"]
        sht.clear()

    sht.range("A1").value = ret_integration
    try:
        sht = wb.sheets.add(name="General", before=None, after=None)
    except:
        sht = wb.sheets["General"]
        sht.clear()

    sht.range("A1").value = ret_general
    wb.save()


if __name__ == '__main__':
    print_hi('PyCharm')
    get_excel_data(gl_fw)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
