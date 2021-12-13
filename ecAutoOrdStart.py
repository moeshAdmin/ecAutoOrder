import ecAutoOrdLib
import ecAutoOrdConfig

def main(source,runType):
    if source!='downloadExport':
        driver = initProcess()

    if source=='hey'||source=='jplabo':
        cookieName = 'cookies_'+source+'.pkl'
        loginURL = ecAutoOrdConfig.GetConfig('loginURL')
        adminURL = ecAutoOrdConfig.GetConfig('adminURL')
        loginChk = ecAutoOrdConfig.GetConfig('loginChk')
        exportURL = ecAutoOrdConfig.GetConfig('exportURL')
        setAccount = ecAutoOrdConfig.GetConfig('account')
        setPassword = ecAutoOrdConfig.GetConfig('password')
        try:
            print('確認登入狀態中')
            ecAutoOrdLib.login(driver,setAccount,setPassword,loginURL,adminURL,cookieName,loginChk,source,runType)
            if runType=='download' or runType=='downloadToday':
                print('開始執行下載程序')
                ecAutoOrdLib.exportData(driver,'cyberbiz',exportURL,runType)

            ecAutoOrdLib.endProcess(driver,runType)
        except Exception as e: 
            ecAutoOrdLib.errorProcess(driver,e,cookieName,source,runType)
                
    elif source=='well':
        cookieName = 'cookies_well.pkl'
        loginURL = ecAutoOrdConfig.GetConfig('loginURL')
        adminURL = ecAutoOrdConfig.GetConfig('adminURL')
        loginChk = ecAutoOrdConfig.GetConfig('loginChk')
        exportURL = ecAutoOrdConfig.GetConfig('exportURL')
        setAccount = ecAutoOrdConfig.GetConfig('account')
        setPassword = ecAutoOrdConfig.GetConfig('password')
        try:
            ecAutoOrdLib.login(driver,setAccount,setPassword,loginURL,adminURL,cookieName,loginChk,source,runType)
            if runType=='download' or runType=='downloadToday':
                print('開始執行下載程序')
                ecAutoOrdLib.exportData(driver,'waca',exportURL,runType)
            ecAutoOrdLib.endProcess(driver,runType)
        except Exception as e: 
            ecAutoOrdLib.errorProcess(driver,e,cookieName,source,runType)

    elif source=='downloadExport':
        print('等待 2 分鐘，待商城端產出訂單')
        #time.sleep(120)
        print('開始擷取已收到下載訂單')
        try:
            ecAutoOrdLib.downloadExportData()
        except Exception as e: 
            print('發生意外錯誤')
            print(e)

#source = 'jplabo'
#runType = 'download'
source = sys.argv[1]
runType = sys.argv[2]

main(source,runType)