import pandas as pd
import bs4 as bs
import requests
import time
import random
from datetime import datetime
from datetime import date
import locale
locale.setlocale(locale.LC_ALL, '')

# Helper Functions


def get_quarters(url):
    resp = requests.get(url)
    soup = bs.BeautifulSoup(resp.text, 'html.parser')
    table = soup.find('table', {'class': 'mctable1'})
    data = []
    if table is None:
        return data
    for row in table.findAll('tr'):
        cols = row.findAll('td')
        cols = [ele.text.strip() for ele in cols]
        data.append([ele for ele in cols if ele])
    x = data[0]
    return x[1:]


def main(filepath, SheetName, month, dateval):
    df = pd.read_excel(filepath, engine='openpyxl')
    df = df.fillna("0")
    urls = df["res_url"].tolist()
    now = datetime.now()
    outdf = pd.DataFrame(columns=[
        "SYMB", f"{month}_R", f"{month}_V", f"{month}_DEPS",
        f"{month}_Interest", f"{month}_Gross NPA", f"{month}_Net NPA", "LUU",
        "RES_CATG_1", "RES_CATG_2", "L_RES_DT", "DATA_AVAILABLE",
        "Last Available Data", "Previous Available Quarter", "TIME STAMP",
        "RES_URL", "SECTOR"
    ])

    # function to check if the value is string or integer and convert it to float in both cases
    def convert(value):
        if isinstance(value, str):
            return float(value.replace(',', ''))
        elif isinstance(value, int):
            return float(value)

    # randomized delay function avoid IP ban
    def delay():
        time.sleep(random.uniform(0, 2))

    # Function to scrap data from the URLs
    def get_data(url):
        delay()
        resp = requests.get(url)
        soup = bs.BeautifulSoup(resp.text, 'html.parser')
        table = soup.find('table', {'class': 'mctable1'})
        data = []
        if table is None:
            return data
        for row in table.findAll('tr'):
            cols = row.findAll('td')
            cols = [ele.text.strip() for ele in cols]
            data.append([ele for ele in cols if ele])
        return data

    for url in urls:
        if df.loc[urls.index(url), "REVENUE"] > 0:
            continue
        data = get_data(url)
        today = date.today()
        if dateval == "NO":
            if df.loc[urls.index(url), "l_res_dt"] < pd.Timestamp(today):
                if len(data) == 0:
                    print(df.loc[urls.index(url), "symb"])
                    outdf.loc[urls.index(url), "RES_URL"] = df.loc[urls.index(
                        url), "res_url"]
                    outdf.loc[urls.index(url), "DATA_AVAILABLE"] = "N"
                    outdf.loc[urls.index(url), "TIME STAMP"] = now.strptime(
                        now.strftime("%Y-%m-%d %H:%M:%S"), "%Y-%m-%d %H:%M:%S")
                    outdf.loc[urls.index(url), "SECTOR"] = df.loc[urls.index(
                        url), "sectr"]
                    outdf.loc[urls.index(
                        url), 'SYMB'] = df.loc[urls.index(url), "symb"]
                    if df.loc[urls.index(url), "LUU"] == "0":
                        outdf.loc[urls.index(url), "LUU"] = "AUTO"
                    else:
                        outdf.loc[urls.index(
                            url), "LUU"] = df.loc[urls.index(url), "LUU"]
                    outdf.loc[urls.index(url), "RES_CATG_1"] = df.loc[urls.index(
                        url), "res_catg1"]
                    outdf.loc[urls.index(url), "RES_CATG_2"] = df.loc[urls.index(
                        url), "res_catg2"]
                    outdf.loc[urls.index(url), "L_RES_DT"] = df.loc[urls.index(
                        url), "l_res_dt"]
                    outdf.loc[urls.index(url), "Last Available Data"] = "NONE"
                    outdf.loc[urls.index(
                        url), "Previous Available Quarter"] = "NONE"
                elif len(data) > 0:
                    dg = pd.DataFrame(data, index=None, columns=None)
                    dg.columns = dg.iloc[0]
                    dg = dg.drop(dg.index[0])
                    dg.replace('--', 0, inplace=True)
                    dg.fillna(0, inplace=True)
                    check = 0
                    b = 0
                    gc = 0
                    for i in dg.columns:
                        if i == month:
                            if (len(dg.columns) > dg.columns.get_loc(i) + 1):
                                prevQ = dg.columns[dg.columns.get_loc(i) + 1]
                                gc = 1
                            dg = dg.take([0, dg.columns.get_loc(i)], axis=1)
                            check = 1
                            break
                    if len(dg.columns) > 1 and check == 1:
                        print(df.loc[urls.index(url), "symb"])
                        outdf.loc[urls.index(
                            url), 'SYMB'] = df.loc[urls.index(url), "symb"]
                        outdf.loc[urls.index(url), "LUU"] = "AUTO"
                        outdf.loc[urls.index(url), "RES_CATG_1"] = df.loc[urls.index(
                            url), "res_catg1"]
                        outdf.loc[urls.index(url), "RES_CATG_2"] = df.loc[urls.index(
                            url), "res_catg2"]
                        outdf.loc[urls.index(url), "L_RES_DT"] = df.loc[urls.index(
                            url), "l_res_dt"]
                        outdf.loc[urls.index(url), "RES_URL"] = df.loc[urls.index(
                            url), "res_url"]
                        outdf.loc[urls.index(url), "SECTOR"] = df.loc[urls.index(
                            url), "sectr"]
                        if gc == 1:
                            outdf.loc[urls.index(
                                url), "Previous Available Quarter"] = prevQ
                        elif gc == 0:
                            outdf.loc[urls.index(
                                url), "Previous Available Quarter"] = "NONE"
                        if dg.loc[1][1] == 0 and (dg.loc[28][1]):
                            outdf.loc[urls.index(url), f"{month}_R"] = 0.01
                            outdf.loc[urls.index(url), f"{month}_V"] = convert(
                                dg.loc[28][1])
                            outdf.loc[urls.index(url), f"{month}_DEPS"] = convert(
                                dg.loc[37][1])
                            outdf.loc[urls.index(url), f"{month}_Interest"] = convert(
                                dg.loc[20][1])
                        elif "bank" in outdf.loc[urls.index(url), "SECTOR"].lower():
                            b = 1
                        elif b == 0:
                            outdf.loc[urls.index(url), f"{month}_R"] = convert(
                                dg.loc[1][1])
                            outdf.loc[urls.index(url), f"{month}_V"] = convert(
                                dg.loc[28][1])
                            outdf.loc[urls.index(url), f"{month}_DEPS"] = convert(
                                dg.loc[37][1])
                        outdf.loc[urls.index(url), f"{month}_Interest"] = convert(
                            dg.loc[20][1])
                        outdf.loc[urls.index(url), "DATA_AVAILABLE"] = "Y"
                        outdf.loc[urls.index(
                            url), "Last Available Data"] = dg.columns[1]
                        outdf.loc[urls.index(url), "TIME STAMP"] = now.strptime(
                            now.strftime("%Y-%m-%d %H:%M:%S"), "%Y-%m-%d %H:%M:%S")
                        if "bank" in outdf.loc[urls.index(url), "SECTOR"].lower():
                            outdf.loc[urls.index(url), f"{month}_R"] = convert(
                                dg.loc[2][1])
                            outdf.loc[urls.index(url), f"{month}_V"] = convert(
                                dg.loc[20][1])
                            outdf.loc[urls.index(url), f"{month}_DEPS"] = convert(
                                dg.loc[29][1])
                            outdf.loc[urls.index(
                                url), f"{month}_Interest"] = 0
                            outdf.loc[urls.index(url), f"{month}_Gross NPA"] = convert(
                                dg.loc[35][1])
                            outdf.loc[urls.index(url), f"{month}_Net NPA"] = convert(
                                dg.loc[36][1])
                    else:
                        outdf.loc[urls.index(url), "RES_URL"] = df.loc[urls.index(
                            url), "res_url"]
                        outdf.loc[urls.index(url), "DATA_AVAILABLE"] = "N"
                        outdf.loc[urls.index(url), "TIME STAMP"] = now.strptime(
                            now.strftime("%Y-%m-%d %H:%M:%S"), "%Y-%m-%d %H:%M:%S")
                        outdf.loc[urls.index(url), "SECTOR"] = df.loc[urls.index(
                            url), "sectr"]
                        outdf.loc[urls.index(
                            url), 'SYMB'] = df.loc[urls.index(url), "symb"]
                        outdf.loc[urls.index(url), "LUU"] = "AUTO"
                        outdf.loc[urls.index(url), "RES_CATG_1"] = df.loc[urls.index(
                            url), "res_catg1"]
                        outdf.loc[urls.index(url), "RES_CATG_2"] = df.loc[urls.index(
                            url), "res_catg2"]
                        outdf.loc[urls.index(url), "L_RES_DT"] = df.loc[urls.index(
                            url), "l_res_dt"]
                        outdf.loc[urls.index(
                            url), "Last Available Data"] = dg.columns[1]
                        outdf.loc[urls.index(
                            url), "Previous Available Quarter"] = "NONE"
            else:
                break
        elif dateval == "YES":
            if len(data) == 0:
                print(df.loc[urls.index(url), "symb"])
                outdf.loc[urls.index(url), "RES_URL"] = df.loc[urls.index(
                    url), "res_url"]
                outdf.loc[urls.index(url), "DATA_AVAILABLE"] = "N"
                outdf.loc[urls.index(url), "TIME STAMP"] = now.strptime(
                    now.strftime("%Y-%m-%d %H:%M:%S"), "%Y-%m-%d %H:%M:%S")
                outdf.loc[urls.index(url), "SECTOR"] = df.loc[urls.index(
                    url), "sectr"]
                outdf.loc[urls.index(
                    url), 'SYMB'] = df.loc[urls.index(url), "symb"]
                if df.loc[urls.index(url), "LUU"] == "0":
                    outdf.loc[urls.index(url), "LUU"] = "AUTO"
                else:
                    outdf.loc[urls.index(
                        url), "LUU"] = df.loc[urls.index(url), "LUU"]
                outdf.loc[urls.index(url), "RES_CATG_1"] = df.loc[urls.index(
                    url), "res_catg1"]
                outdf.loc[urls.index(url), "RES_CATG_2"] = df.loc[urls.index(
                    url), "res_catg2"]
                outdf.loc[urls.index(url), "L_RES_DT"] = df.loc[urls.index(
                    url), "l_res_dt"]
                outdf.loc[urls.index(url), "Last Available Data"] = "NONE"
                outdf.loc[urls.index(
                    url), "Previous Available Quarter"] = "NONE"
            elif len(data) > 0:
                dg = pd.DataFrame(data, index=None, columns=None)
                dg.columns = dg.iloc[0]
                dg = dg.drop(dg.index[0])
                dg.replace('--', 0, inplace=True)
                dg.fillna(0, inplace=True)
                check = 0
                b = 0
                gc = 0
                for i in dg.columns:
                    if i == month:
                        if (len(dg.columns) > dg.columns.get_loc(i) + 1):
                            prevQ = dg.columns[dg.columns.get_loc(i) + 1]
                            gc = 1
                        dg = dg.take([0, dg.columns.get_loc(i)], axis=1)
                        check = 1
                        break
                if len(dg.columns) > 1 and check == 1:
                    print(df.loc[urls.index(url), "symb"])
                    outdf.loc[urls.index(
                        url), 'SYMB'] = df.loc[urls.index(url), "symb"]
                    outdf.loc[urls.index(url), "LUU"] = "AUTO"
                    outdf.loc[urls.index(url), "RES_CATG_1"] = df.loc[urls.index(
                        url), "res_catg1"]
                    outdf.loc[urls.index(url), "RES_CATG_2"] = df.loc[urls.index(
                        url), "res_catg2"]
                    outdf.loc[urls.index(url), "L_RES_DT"] = df.loc[urls.index(
                        url), "l_res_dt"]
                    outdf.loc[urls.index(url), "RES_URL"] = df.loc[urls.index(
                        url), "res_url"]
                    outdf.loc[urls.index(url), "SECTOR"] = df.loc[urls.index(
                        url), "sectr"]
                    if gc == 1:
                        outdf.loc[urls.index(
                            url), "Previous Available Quarter"] = prevQ
                    elif gc == 0:
                        outdf.loc[urls.index(
                            url), "Previous Available Quarter"] = "NONE"
                    if dg.loc[1][1] == 0 and (dg.loc[28][1]):
                        outdf.loc[urls.index(url), f"{month}_R"] = 0.01
                        outdf.loc[urls.index(url), f"{month}_V"] = convert(
                            dg.loc[28][1])
                        outdf.loc[urls.index(url), f"{month}_DEPS"] = convert(
                            dg.loc[37][1])
                        outdf.loc[urls.index(url), f"{month}_Interest"] = convert(
                            dg.loc[20][1])
                    elif "bank" in outdf.loc[urls.index(url), "SECTOR"].lower():
                        b = 1
                    elif b == 0:
                        outdf.loc[urls.index(url), f"{month}_R"] = convert(
                            dg.loc[1][1])
                        outdf.loc[urls.index(url), f"{month}_V"] = convert(
                            dg.loc[28][1])
                        outdf.loc[urls.index(url), f"{month}_DEPS"] = convert(
                            dg.loc[37][1])
                    outdf.loc[urls.index(url), f"{month}_Interest"] = convert(
                        dg.loc[20][1])
                    outdf.loc[urls.index(url), "DATA_AVAILABLE"] = "Y"
                    outdf.loc[urls.index(
                        url), "Last Available Data"] = dg.columns[1]
                    outdf.loc[urls.index(url), "TIME STAMP"] = now.strptime(
                        now.strftime("%Y-%m-%d %H:%M:%S"), "%Y-%m-%d %H:%M:%S")
                    if "bank" in outdf.loc[urls.index(url), "SECTOR"].lower():
                        outdf.loc[urls.index(url), f"{month}_R"] = convert(
                            dg.loc[2][1])
                        outdf.loc[urls.index(url), f"{month}_V"] = convert(
                            dg.loc[20][1])
                        outdf.loc[urls.index(url), f"{month}_DEPS"] = convert(
                            dg.loc[29][1])
                        outdf.loc[urls.index(
                            url), f"{month}_Interest"] = 0
                        outdf.loc[urls.index(url), f"{month}_Gross NPA"] = convert(
                            dg.loc[35][1])
                        outdf.loc[urls.index(url), f"{month}_Net NPA"] = convert(
                            dg.loc[36][1])
                else:
                    outdf.loc[urls.index(url), "RES_URL"] = df.loc[urls.index(
                        url), "res_url"]
                    outdf.loc[urls.index(url), "DATA_AVAILABLE"] = "N"
                    outdf.loc[urls.index(url), "TIME STAMP"] = now.strptime(
                        now.strftime("%Y-%m-%d %H:%M:%S"), "%Y-%m-%d %H:%M:%S")
                    outdf.loc[urls.index(url), "SECTOR"] = df.loc[urls.index(
                        url), "sectr"]
                    outdf.loc[urls.index(
                        url), 'SYMB'] = df.loc[urls.index(url), "symb"]
                    outdf.loc[urls.index(url), "LUU"] = "AUTO"
                    outdf.loc[urls.index(url), "RES_CATG_1"] = df.loc[urls.index(
                        url), "res_catg1"]
                    outdf.loc[urls.index(url), "RES_CATG_2"] = df.loc[urls.index(
                        url), "res_catg2"]
                    outdf.loc[urls.index(url), "L_RES_DT"] = df.loc[urls.index(
                        url), "l_res_dt"]
                    outdf.loc[urls.index(
                        url), "Last Available Data"] = dg.columns[1]
                    outdf.loc[urls.index(
                        url), "Previous Available Quarter"] = "NONE"

    outdf.replace('--', 0, inplace=True)
    outdf.fillna(0, inplace=True)
    current_time = time.strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"Qtr_res_out_{month}_{current_time}.xlsx"
    outdf.to_excel(filename, index=False, sheet_name=SheetName)
    print("File saved as: ", filename)
