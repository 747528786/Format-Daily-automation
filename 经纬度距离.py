import requests
import pandas as pd
import numpy as np
import time

df = pd.read_excel('新建 Microsoft Excel 工作表.xlsx')
list = df.values.tolist()
list1 = []
for i in range(len(list)):
    list1.append(list[i][0])

def main():
    start = time.time()

    list_info = change_addresss(addresslist=list1, key="高德秘钥")

    geodistance(list=list_info).to_excel(excel_writer="距离.xlsx", index=False)

    end = time.time()
    print('运行时间: %s秒' % (end - start))


def change_addresss(addresslist,key):
    base = "https://restapi.amap.com/v3/geocode/geo?"
    output = 'json'
    batch = 'true'
    for i in range(0,len(addresslist)):
        address = addresslist[i]
        currentUrl = base + "output=" + output + "&batch=" + batch + "&address=" + address + "&key=" + key
        response = requests.get(currentUrl)
        answer = response.json()
        print(answer)
        if answer['status'] == '1' and answer['info'] == 'OK' and answer['count'] == '1':
            tmpList = answer['geocodes']
            formatted_address = tmpList[0]['formatted_address']
            coordString = tmpList[0]['location']
            coordList = coordString.split(',')
            list[i].append(formatted_address)
            list[i].append(float(coordList[0]))
            list[i].append(float(coordList[1]))
        elif answer['status'] == '0':
            list[i].extend(["接口调用失败", 0, 0])
        else:
            list[i].extend(['地址转换失败', 0, 0])
    return list

def geodistance(list):
    df = pd.DataFrame(list, columns=["地址", "lng1", "lat1", "格式后地址", "lng2", "lat2"])
    arr_lng1 = np.array(df["lng1"].apply(lambda x: x).tolist())
    arr_lat1 = np.array(df["lat1"].apply(lambda x: x).tolist())
    arr_lng2 = np.array(df["lng2"].apply(lambda x: x).tolist())
    arr_lat2 = np.array(df["lat2"].apply(lambda x: x).tolist())
    r_lng1 = np.radians(arr_lng1)
    r_lat1 = np.radians(arr_lat1)
    r_lng2 = np.radians(arr_lng2)
    r_lat2 = np.radians(arr_lat2)
    d_lng = r_lng2 - r_lng1
    d_lat = r_lat2 - r_lat1
    a = np.sin(d_lat / 2) ** 2 + np.cos(arr_lat1) * np.cos(arr_lat2) * np.sin(d_lng / 2) ** 2
    distance = 2 * np.arcsin(np.sqrt(a)) * 6371 * 1000  # 地球平均半径，6371km
    distance = np.round(distance, 2)
    df["距离"] = distance.tolist()

    return df




if __name__ == '__main__':
    main()

