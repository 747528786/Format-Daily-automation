import json
import pandas as pd
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.ocr.v20181119 import ocr_client, models

df = pd.read_excel("OCR.xlsx")
list1 = df.values.tolist()
print(list1)


def ocr_license(imageUrlList):
    for i in range(0, len(imageUrlList)):
        try:
            cred = credential.Credential("AKIDKRLyLRGo56qJVQ4RXN3WidIgndCUVGkx", "yf5dvwf8LGuIPFo67rKWS8XfW6OP8bnh")
            httpProfile = HttpProfile()
            httpProfile.endpoint = "ocr.tencentcloudapi.com"

            clientProfile = ClientProfile()
            clientProfile.httpProfile = httpProfile
            client = ocr_client.OcrClient(cred, "ap-beijing", clientProfile)

            req = models.BizLicenseOCRRequest()
            params = {
                "ImageUrl": imageUrlList[i][0]

            }
            print(params)
            req.from_json_string(json.dumps(params))

            resp = client.BizLicenseOCR(req)
            a = json.loads(resp.to_json_string())["RegNum"]
            print(a)
            list1[i].append(a)


        except TencentCloudSDKException as err:
            list1[i].append(str(err))
            print(err)

    return list1


if __name__ == '__main__':
    ocr_license(imageUrlList=list1)
    df1 = pd.DataFrame(list1)
    df1.to_excel(excel_writer="OCR结果.xlsx", index=False)
