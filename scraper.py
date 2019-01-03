import requests
from bs4 import BeautifulSoup
import json
import csv
import os,sys,datetime, time
import pandas as pd
global RadScriptManager1_TSM, RadStyleSheetManager1_TSSM, __VIEWSTATEFIELDCOUNT, __VIEWSTATES,__VIEWSTATEGENERATOR,__EVENTVALIDATION
global searchKeys, product
global policyVals, policyKeys
"""define static variables"""

userName = "eligra1"
passWord = "VR8469"
policyUrl = "https://policy.velocityrisk.com/Policy/Policy.aspx"
header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36','X-MicrosoftAjax': 'Delta=true', 'X-Requested-With': 'XMLHttpRequest'}

""" Definition of Functions """
def getElementbySpan(soup,span,num=1):
    try:
        eles = soup.find("span",string=span).parent.parent.select('td')
        if eles[num].findChild('input',{'type':'text'}) is None:data = eles[num].findChild('option',{'selected':'selected'}).text
        else:data = eles[num].findChild('input',{'type':'text'}).get('value')
    except:data = ""
    if data==None:data = ""
    return data

def getElementbySpan2(soup, span):
    data = ""
    try:
        eleId = soup.find("span", string=span).parent.get('id')
        no = int(eleId.split("TableCell")[1])
        eles = soup.find("span",string=span).parent.parent.findChildren()
        for child in eles:
            try:
                id = child.get('id')
                if int(id.split("TableCell")[1]) > no :
                    if child.findChild('input', {'type': 'text'}) is None:data = child.findChild('option', {'selected': 'selected'}).text
                    else:data = child.findChild('input', {'type': 'text'}).get('value')
                    break
            except:
                continue
    except:data = ""
    if data==None :data = ""
    return data

def getElementbyLabel(soup, label):
    try:
        ele = soup.find("label",text=label).parent
        if ele.findChild('input',{'type':'text'}) is None and ele.findChild('select') is None: return -1
        if ele.findChild('input',{'type':'text'}) is None:data = ele.findChild('option',{'selected':'selected'}).text
        else:
            data = ele.findChild('input',{'type':'text'}).get('value')
            if data=="" or data==None:data = ele.findChild('input', {'type': 'hidden'}).get('value')
    except:data = ""
    if data == None :data = ""
    return data

def getElementbyLabel2(soup,label):
    try:
        ele = soup.find("label",text=label).parent.parent
        if ele.findChild('input',{'type':'text'}) is None and ele.findChild('select') is None: return -1
        if ele.findChild('input',{'type':'text'}) is None:data = ele.findChild('option',{'selected':'selected'}).text
        else:data = ele.findChild('input',{'type':'text'}).get('value')
    except:data = ""
    if data == None :data = ""
    return data

def getJsonFromTable(trs):
    arr = []
    values = []
    keys = []
    for th in trs[0]:
        try:
            if (th.text!="\u00a0" and th != "\n"):
                keys.append(th.text)
        except:
            pass
    cnt = 0
    for tr in trs:
        if cnt > 0:
            if dict(tr.attrs)["class"][0] != "rgGroupHeader":
                for td in tr:
                    if (td != "\n" and td.text !='\u00a0' and td.text !='No records to display.'):
                        values.append(td.text)
                unit = dict(zip(keys, values))
                arr.append(unit)
                values = []
        cnt+=1
    return json.dumps(arr)

def getEle(soup, tag, idValue):
    try:
        if tag=='select':data = soup.find(tag,{'id':idValue}).find('option',{'selected':'selected'}).text
        else:data = soup.find(tag,{'id':idValue}).get('value')
    except:data = ""
    if data == None or data.replace('\n',"").replace(" ","") == "":data = ""
    return data

def getCleanLabel(label):
    return label.replace('\t','').replace('\n','').replace('\r','')

def getCookie(soup):
    global RadScriptManager1_TSM, RadStyleSheetManager1_TSSM, __VIEWSTATEFIELDCOUNT, __VIEWSTATES,__VIEWSTATEGENERATOR,__EVENTVALIDATION
    try:
        RadScriptManager1_TSM = soup.find('input', {'id': 'RadScriptManager1_TSM'}).get('value')
        RadStyleSheetManager1_TSSM = soup.find('input', {'id': 'RadStyleSheetManager1_TSSM'}).get('value')
        __VIEWSTATEFIELDCOUNT = soup.find('input', {'id': '__VIEWSTATEFIELDCOUNT'}).get('value')
        __VIEWSTATES = {}
        for i in range(int(__VIEWSTATEFIELDCOUNT)):
            if i == 0:
                try:
                    val = soup.find('input', {'id': '__VIEWSTATE'}).get('value')
                    __VIEWSTATES.update({'__VIEWSTATE': val})
                except:
                    pass
                continue
            try:
                id = "__VIEWSTATE" + str(i)
                val = soup.find(id=id).get('value')
                __VIEWSTATES.update({id: val})
            except:
                pass
        __VIEWSTATEGENERATOR = soup.find('input', {'id': '__VIEWSTATEGENERATOR'}).get('value')
        __EVENTVALIDATION = soup.find('input', {'id': '__EVENTVALIDATION'}).get('value')
    except:
        pass

def outputCSV(fileName, heads, bodys):
    if not os.path.isfile(fileName):
        with open(fileName, "a") as output:
            writer = csv.writer(output, lineterminator='\n', quotechar='"')
            writer.writerow(heads)
            writer.writerow(bodys)
    else:
        with open(fileName, "a") as output:
            writer = csv.writer(output, lineterminator='\n', quotechar='"')
            writer.writerow(bodys)

def getPostData():
    postData = {
        'ctl00$RadScriptManager1': 'ctl00$ctl00$ContentPlaceHolderBody$trsPolicyPanel|ctl00$ContentPlaceHolderBody$trsPolicy',
        'RadScriptManager1_TSM': RadScriptManager1_TSM, 'RadStyleSheetManager1_TSSM': RadStyleSheetManager1_TSSM,
        '__EVENTTARGET': 'ctl00$ContentPlaceHolderBody$trsPolicy',
        '__VIEWSTATEFIELDCOUNT': __VIEWSTATEFIELDCOUNT, '__VIEWSTATEGENERATOR': __VIEWSTATEGENERATOR,
    }
    for i in range(int(__VIEWSTATEFIELDCOUNT)):
        if i == 0:
            postData.update({'__VIEWSTATE': __VIEWSTATES.get("__VIEWSTATE")})
        else:
            postData.update({'__VIEWSTATE' + str(i): __VIEWSTATES.get("__VIEWSTATE" + str(i))})
    return postData

def login():
    #define Session
    session = requests.session()
    """get token and other fields in login page"""
    Homeurl = "https://policy.velocityrisk.com/login.aspx?ReturnUrl=%2f"
    res = session.get(Homeurl)
    soup = BeautifulSoup(res.text, 'html.parser')
    try:
        viewState           = soup.find('input', {'id': '__VIEWSTATE'}).get('value')
        viewStateGenerator  = soup.find('input', {'id': '__VIEWSTATEGENERATOR'}).get('value')
        viewEventValidation = soup.find('input', {'id': '__EVENTVALIDATION'}).get('value')
    except:pass

    """login Action & get Params"""
    loginUrl = "https://policy.velocityrisk.com/login.aspx?ReturnUrl=%2f"
    postData = {'AquantLogin$UserName': userName,'AquantLogin$Password': passWord,
                'AquantLogin$LoginButton': 'Log In','__VIEWSTATE': viewState,
                '__VIEWSTATEGENERATOR': viewStateGenerator,'__EVENTVALIDATION' : viewEventValidation}

    res = session.post( loginUrl, data=postData, allow_redirects=False)
    soup = BeautifulSoup(res.text, "html.parser")
    atag = soup.find('a').get('href')
    if (atag == "/TermsOfUse/TermsOfUse.aspx"):
        url ="https://policy.velocityrisk.com/TermsOfUse/TermsOfUse.aspx"
        postData = { "__VIEWSTATE" : viewState,"__VIEWSTATEGENERATOR" : viewStateGenerator,
                     "__EVENTVALIDATION":viewEventValidation,"btnAccept" : "Accept" }
        session.post(url,data=postData)
        res = session.get("https://policy.velocityrisk.com/Home.aspx")
    else:
        url = "https://policy.velocityrisk.com/Home.aspx"
        res = session.get(url)
    soup = BeautifulSoup(res.text, "html.parser")
    getCookie(soup)
    return session

def excel_download():
    session = login()
    """get policy key excel file"""
    url = "https://policy.velocityrisk.com/Customer/CustomerSearch.aspx"
    res = session.get(url)
    soup = BeautifulSoup(res.text, "html.parser")
    getCookie(soup)

    postData = {
        'RadScriptManager1_TSM': RadScriptManager1_TSM,
        'RadStyleSheetManager1_TSSM': RadStyleSheetManager1_TSSM,
        '__VIEWSTATEFIELDCOUNT': __VIEWSTATEFIELDCOUNT,
        '__VIEWSTATEGENERATOR': __VIEWSTATEGENERATOR,
        '__EVENTVALIDATION': __EVENTVALIDATION,
        'ctl00$ContentPlaceHolderBody$CustomerSearch1$txtPhone': '',
        'ctl00$ContentPlaceHolderBody$CustomerSearch1$ddlPolicyStatusFilter': '8',
        'ctl00$ContentPlaceHolderBody$CustomerSearch1$btnXLSReport.x': '10',
        'ctl00$ContentPlaceHolderBody$CustomerSearch1$btnXLSReport.y': '18',
        'ctl00$ContentPlaceHolderBody$CustomerSearch1$RadGridCustomerSearch$ctl00$ctl03$ctl01$PageSizeComboBox': '40'
    }
    for i in range(int(__VIEWSTATEFIELDCOUNT)):
        if i == 0:
            postData.update({'__VIEWSTATE': __VIEWSTATES.get("__VIEWSTATE")})
        else:
            postData.update({'__VIEWSTATE' + str(i): __VIEWSTATES.get("__VIEWSTATE" + str(i))})
    header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded'}
    res = session.post(url,data=postData,headers = header)
    with open('policykey.xls', 'wb') as file:
        file.write(res.content)
    file.close()

def gotoActionItemsPage(session,searchKey):
    ActionUrl = "https://policy.velocityrisk.com/HomeItems/ActionItems.aspx"
    res = session.get(ActionUrl)
    soup = BeautifulSoup(res.text, "html.parser")
    getCookie(soup)
    postData = {
        'RadScriptManager1_TSM': RadScriptManager1_TSM,
        'RadStyleSheetManager1_TSSM': RadStyleSheetManager1_TSSM,
        '__VIEWSTATEFIELDCOUNT': '2',
        '__VIEWSTATEGENERATOR': __VIEWSTATEGENERATOR,
        '__EVENTVALIDATION': __EVENTVALIDATION,
        'ctl00$ContentPlaceHolderBody$TopNavOne1$txtPolicyIdWithprefix': searchKey,
        'ctl00$ContentPlaceHolderBody$TopNavOne1$btnGoId': 'GO',
        'ctl00$ContentPlaceHolderBody$ActionItems1$ddlActionItemType': 'ALL',
        'ctl00$ContentPlaceHolderBody$ActionItems1$ddlAgencyProducer': 'Producer',
        'ctl00$ContentPlaceHolderBody$ActionItems1$ddlPolicyStatus': 'ALL',
        'ctl00$ContentPlaceHolderBody$ActionItems1$HiddenSource': 'PENDING',
    }


    for i in range(int(__VIEWSTATEFIELDCOUNT)):
        if i == 0:postData.update({'__VIEWSTATE': __VIEWSTATES.get("__VIEWSTATE")})
        else:postData.update({'__VIEWSTATE' + str(i): __VIEWSTATES.get("__VIEWSTATE" + str(i))})
    res = session.post(ActionUrl, data=postData)
    soup = BeautifulSoup(res.text, "html.parser")
    getCookie(soup)
    return session

def getPolicyBilling(session,searchKey):
    global policyVals, policyKeys, policyUrl, header
    """get policy data & get Hidden fields & Policy History data"""
    ActionUrl = "https://policy.velocityrisk.com/HomeItems/ActionItems.aspx"
    postData = {
        'RadScriptManager1_TSM': RadScriptManager1_TSM,
        'RadStyleSheetManager1_TSSM': RadStyleSheetManager1_TSSM,
        '__VIEWSTATEFIELDCOUNT': __VIEWSTATEFIELDCOUNT,
        '__VIEWSTATEGENERATOR': __VIEWSTATEGENERATOR,
        '__EVENTVALIDATION': __EVENTVALIDATION,
        'ctl00$ContentPlaceHolderBody$TopNavOne1$txtPolicyIdWithprefix': searchKey,
        'ctl00$ContentPlaceHolderBody$TopNavOne1$btnGoId': 'GO',
        'ctl00$ContentPlaceHolderBody$ActionItems1$ddlActionItemType': 'ALL',
        'ctl00$ContentPlaceHolderBody$ActionItems1$ddlAgencyProducer': 'Producer',
        'ctl00$ContentPlaceHolderBody$ActionItems1$ddlPolicyStatus': 'ALL',
        'ctl00$ContentPlaceHolderBody$ActionItems1$HiddenSource': 'PENDING'
    }
    for i in range(int(__VIEWSTATEFIELDCOUNT)):
        if i == 0:postData.update({'__VIEWSTATE': __VIEWSTATES.get("__VIEWSTATE")})
        else:postData.update({'__VIEWSTATE' + str(i): __VIEWSTATES.get("__VIEWSTATE" + str(i))})
    res = session.post(ActionUrl, data=postData)
    soup = BeautifulSoup(res.text, "html.parser")
    getCookie(soup)

    """get Policy History/Billing"""
    policyVals = []
    policyKeys = [
        'Renewal Payor',
        'Pay Plan',
        'Balance',
        'Action Items(JSON)',
        'Transaction ID',
        'Tran Cd',
        'Created Date',
        'Transaction Eff Date',
        'Direct Prem',
        'Fees',
        'Total Direct',
        'Accounting Detail(JSON)',
        'Memos(JSON)'
    ]
    try:renewalPayor = getElementbySpan(soup, "Renewal Payor")
    except:renewalPayor = ""
    try:payPlan = soup.find('select', {'id': 'ContentPlaceHolderBody_PolicyHistoryBillingBody_ddlPaymentPlans'}).find('option', {'selected': 'selected'}).text
    except:payPlan = ""
    try:payPlanNum = soup.find('select',{'id': 'ContentPlaceHolderBody_PolicyHistoryBillingBody_ddlPaymentPlans'}).find('option',{'selected': 'selected'}).get('value')
    except:payPlanNum = "0"
    try:balance = soup.find('span', {'id': 'ContentPlaceHolderBody_PolicyHistoryBillingBody_lblBalance'}).text
    except:balance = ""
    try:actionItems = getJsonFromTable(soup.find('table', id='ContentPlaceHolderBody_PolicyHistoryBillingBody_dgActionItems_ctl00').find_all('tr'))
    except:actionItems = "[{}]"
    try:transHistory = getJsonFromTable(soup.find('table', {'id': 'ContentPlaceHolderBody_PolicyHistoryBillingBody_RadGridTranHistory_ctl00'}).find_all('tr'))
    except:transHistory = "[{}]"
    tranhis = json.loads(transHistory)
    if len(tranhis[0]) == 0:
        tranID = ""
        tranCd = ""
        createdDate = ""
        tranEffDate = ""
        directPrem = ""
        fees = ""
        totalDirect = ""
    else:
        tranID = tranhis[0]['Transaction Id']
        tranCd = tranhis[0]['Tran Cd']
        createdDate = tranhis[0]['Created Date']
        tranEffDate = tranhis[0]['Transaction Eff Date']
        directPrem = tranhis[0]['Direct Prem']
        fees = tranhis[0]['Fees']
        totalDirect = tranhis[0]['Total Direct']
    try:accountingDetail = getJsonFromTable(soup.find('table', {'id': 'ContentPlaceHolderBody_PolicyHistoryBillingBody_RadGridAccountingDetail_ctl00'}).find_all('tr'))
    except:accountingDetail = "[{}]"
    policyVals.append(renewalPayor)
    policyVals.append(payPlan)
    policyVals.append(balance)
    policyVals.append(actionItems)
    policyVals.append(tranID)
    policyVals.append(tranCd)
    policyVals.append(createdDate)
    policyVals.append(tranEffDate)
    policyVals.append(directPrem)
    policyVals.append(fees)
    policyVals.append(totalDirect)
    policyVals.append(accountingDetail)
    """get Memos"""
    memos = []
    postData = getPostData()
    postData['ctl00$RadScriptManager1'] = 'ctl00$ctl00$ContentPlaceHolderFooter$PolicyFooter1$btnMemosPendingPanel|ctl00$ContentPlaceHolderFooter$PolicyFooter1$btnMemosPending'
    postData['__EVENTTARGET'] = 'ctl00$ContentPlaceHolderFooter$PolicyFooter1$btnMemosPending'
    res = session.get("https://policy.velocityrisk.com/Memos.aspx", headers=header)
    soup = BeautifulSoup(res.text, "html.parser")

    memoDiv = soup.find('div', {'id': 'Memos1_RadGridMemos_GridData'})
    trs = memoDiv.find_all('tr')
    for tr in trs:
        try:
            created = tr.select('td')[0].get_text()
            text = tr.select('td')[1].get_text()
            createdby = tr.select('td')[2].get_text()
            memo = {
                "Created": created,
                "Text": text.replace('\n', '').replace('\r', ''),
                "Created By": createdby
            }
            memos.append(memo)
        except:pass
    policyVals.append(json.dumps(memos))
    return session

def getApplicant(session):
    global product, header, policyUrl
    global RadScriptManager1_TSM,RadStyleSheetManager1_TSSM,__VIEWSTATEFIELDCOUNT,__VIEWSTATEGENERATOR
    """get aplicant"""
    postData = getPostData()
    postData['__EVENTARGUMENT'] = '{"type":0,"index":"0"}'
    postData['ContentPlaceHolderBody_trsPolicy_ClientState'] = '{"selectedIndexes":["0"],"logEntries":[],"scrollState":{}}'
    postData['ContentPlaceHolderBody_rmpPolicy_ClientState'] = '{"selectedIndex":0,"changeLog":[]}'

    res = session.post(policyUrl, data=postData, headers=header)
    soup = BeautifulSoup(res.text, "html.parser")

    applicantKeys = [
        'Effective Date',
        'Product',
        'Applicant Insured Type',
        'Applicant First',
        'Applicant M.I.',
        'Applicant Last (or Company)',
        'Applicant Prior Address',
        'Applicant Prior Address Line 2',
        'Applicant Prior Address City',
        'Applicant Prior Address State',
        'Applicant Prior Address Zip',
        'Applicant Prior Address Zip2',
        'Applicant International Address',
        'Applicant Mailing Address',
        'Applicant Mailing Address Line 2',
        'Applicant Mailing Address City',
        'Applicant Mailing Address State',
        'Applicant Mailing Address Zip',
        'Applicant Mailing Address Zip2',
        'Applicant Date of Birth',
        'Applicant Primary Phone',
        'Applicant Cell',
        'Applicant E-mail',
        'Applicant Marital Status',
        'Applicant Children in Household'
    ]
    applicantVals = []
    applicantTag = soup.find("table", id="ContentPlaceHolderBody_ApplicantMailingBody_tblPolicyApplicant")
    applicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_dtEffectiveDate'))
    product = getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtProductCd')
    applicantVals.append(product)
    try:applicantVals.append(getElementbySpan(applicantTag, "Insured Type"))
    except:applicantVals.append("")
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsFirstName'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsMiddleInitial'))  # MI
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsLastName'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsPrimaryAddress1'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsPrimaryAddress2'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsPrimaryCity'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsPrimaryState'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsPrimaryZipCd'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsPrimaryZip4'))
    applicantInternationalAddress = "", applicantVals.append("")
    applicantVals.append(getEle(applicantTag, 'input','ContentPlaceHolderBody_ApplicantMailingBody_txtInsMailingAddress1'))  # applicantMailingAddress
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsMailingAddress2'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsMailingCity'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsMailingState'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsMailingZipCd'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsMailingZip4'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_RadDateInputInsBirthDt'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsHomePhone'))
    # applicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsWorkPhone'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsCellPhone'))
    applicantVals.append(getEle(applicantTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtInsEmail'))
    applicantVals.append(getEle(applicantTag, 'select', 'ContentPlaceHolderBody_ApplicantMailingBody_MaritalStatus'))
    applicantVals.append(getEle(applicantTag, 'select', 'ContentPlaceHolderBody_ApplicantMailingBody_ChildrenInHousehold'))

    coApplicantKeys = [
        'Co-Applicant Insured Type',
        'Co-Applicant First',
        'Co-Applicant M.I.',
        'Co-Applicant Last (or Company)',
        'Co-Applicant Prior Address',
        'Co-Applicant Prior Address Line 2',
        'Co-Applicant Prior Address City',
        'Co-Applicant Prior Address State',
        'Co-Applicant Prior Address Zip',
        'Co-Applicant Prior Address Zip2',
        'Co-Applicant International Address',
        'Co-Applicant Mailing Address',
        'Co-Applicant Mailing Address Line 2',
        'Co-Applicant Mailing Address City',
        'Co-Applicant Mailing Address State',
        'Co-Applicant Mailing Address Zip',
        'Co-Applicant Mailing Address Zip2',
        'Co-Applicant Date of Birth',
        'Co-Applicant Primary Phone',
        'Co-Applicant Cell',
        'Co-Applicant E-mail'
    ]
    coApplicantVals = []
    coApplicantVals.append(getElementbySpan(soup, "Insured Type"))
    # coApplicantVals.append(getEle(soup, 'select', 'ContentPlaceHolderBody_ApplicantMailingBody_ddlInsuredTypeCoApp'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoFirstName'))  # first Name
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoMiddleInitial'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoLastName'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoPrimaryAddress1'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoPrimaryAddress2'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoPrimaryCity'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoPrimaryState'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoPrimaryZipCd'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoPrimaryZip4'))
    coApplicantInternationalAddress = "", coApplicantVals.append("")
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoMailingAddress1'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoMailingAddress2'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoMailingCity'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoMailingState'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoMailingZipCd'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoMailingZip4'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_RadDateInputCoBirthDt'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoHomePhone'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoCellPhone'))
    coApplicantVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCoEmail'))

    propertyAddressKeys = [
        'Property Address',
        'Property Address Line 2',
        'Property Address City',
        'Property Address State',
        'Property Address Zip',
        'Property Address Zip2',
        'County',
        'Census Block Group'
    ]
    propertyAddressVals = []
    propertyAddressTag = soup.find(id="ContentPlaceHolderBody_ApplicantMailingBody_tblpropertyaddress")
    propertyAddressVals.append(getEle(propertyAddressTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtPropertyAddress1'))
    propertyAddressVals.append(getEle(propertyAddressTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtPropertyAddress2'))
    propertyAddressVals.append(getEle(propertyAddressTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtPropertyAddressCity'))
    propertyAddressVals.append(getEle(propertyAddressTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtPropertyAddressState'))
    propertyAddressVals.append(getEle(propertyAddressTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtPropertyAddressZipCd1'))
    propertyAddressVals.append(getEle(propertyAddressTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtPropertyAddressZipCd2'))
    propertyAddressVals.append(getElementbySpan(propertyAddressTag, "County"))
    propertyAddressVals.append(getEle(propertyAddressTag, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtCensusBlockGroup'))

    fstMortageeKeys = [
        '1st Mortgagee Name',
        '1st Mortgagee International Address',
        '1st Mortgagee Address',
        '1st Mortgagee Address Line 2',
        '1st Mortgagee Address City',
        '1st Mortgagee Address State',
        '1st Mortgagee Address Zip',
        '1st Mortgagee Address Zip2',
        '1st Mortgagee Loan  #',
        '1st Mortgagee Phone  #'
    ]
    fstMortageeVals = []
    fstMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee1Name'))
    fstMortgageeInternationalAddress = "", fstMortageeVals.append("")
    fstMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee1MaillingAdd'))
    fstMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee1MaillingAdd2'))
    fstMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtmortgagee1City'))
    fstMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_TxtMortgagee1State'))
    fstMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee1ZipCd1'))
    fstMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee1ZipCd2'))
    fstMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtmortgagee1Loan'))
    fstMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee1Phone'))

    sndMortageeKeys = [
        '2nd Mortgagee Name',
        '2nd Mortgagee International Address',
        '2nd Mortgagee Address',
        '2nd Mortgagee Address Line 2',
        '2nd Mortgagee Address City',
        '2nd Mortgagee Address State',
        '2nd Mortgagee Address Zip',
        '2nd Mortgagee Address Zip2',
        '2nd Mortgagee Loan  #',
        '2nd Mortgagee Phone  #'
    ]
    sndMortageeVals = []
    sndMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtmortgagee2Name'))
    sndMortgageeInternationalAddress = "", sndMortageeVals.append("")
    sndMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtmortgagee2MaillingAdd'))
    sndMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtmortgagee2MailingAdd2'))
    sndMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtmortgagee2City'))
    sndMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee2State'))
    sndMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee2ZipCd1'))
    sndMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee2ZipCd2'))
    sndMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtmortgagee2Loan'))
    sndMortageeVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtMortgagee2Phone'))

    fstAdditionalInsuredKeys = [
        '1st Additional Insured Insured Type',
        '1st Additional Insured First',
        '1st Additional Insured M.I.',
        '1st Additional Insured Last( or Company)',
        '1st Additional Insured International Address',
        '1st Additional Insured Mailing Address',
        '1st Additional Insured Mailing Address Line 2',
        '1st Additional Insured Mailing Address City',
        '1st Additional Insured Mailing Address State',
        '1st Additional Insured Mailing Address Zip',
        '1st Additional Insured Mailing Address Zip2',
        '1st Additional Insured Primary Phone',
        '1st Additional Insured E-mail'
    ]
    fstAdditionalInsuredVals = []
    fstAdditionalInsuredTag = soup.find(id="ContentPlaceHolderBody_ApplicantMailingBody_Table9")
    fstAdditionalInsuredVals.append(getElementbySpan(fstAdditionalInsuredTag, "Insured Type"))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1Name'))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1MiddleName'))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1LastName'))
    fstAdditionalInsuredInternationalAddress = "", fstAdditionalInsuredVals.append("")
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1maillingAdd'))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1maillingAdd2'))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1City'))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1State'))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1ZipCd1'))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1ZipCd2'))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1Phone'))
    fstAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured1Email'))

    sndAdditionalInsuredKeys = [
        '2nd Additional Insured Insured Type',
        '2nd Additional Insured First',
        '2nd Additional Insured M.I.',
        '2nd Additional Insured Last (or Company)',
        '2nd Additional Insured International Address',
        '2nd Additional Insured Mailing Address',
        '2nd Additional Insured Mailing Address Line 2',
        '2nd Additional Insured Mailing Address City',
        '2nd Additional Insured Mailing Address State',
        '2nd Additional Insured Mailing Address Zip',
        '2nd Additional Insured Mailing Address Zip2',
        '2nd Additional Insured Primary Phone',
        '2nd Additional Insured E-mail'
    ]
    sndAdditionalInsuredVals = []
    sndAdditionalInsuredTag = soup.find(id="ContentPlaceHolderBody_ApplicantMailingBody_Table12")
    sndAdditionalInsuredVals.append(getElementbySpan(sndAdditionalInsuredTag, "Insured Type"))
    # sndAdditionalInsuredVals.append(getEle(soup,'select','ContentPlaceHolderBody_ApplicantMailingBody_ddlAdd2InsuredType'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2Name'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2MiddleName'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2LastName'))
    sndAdditionalInsuredInternationalAddress = "", sndAdditionalInsuredVals.append("")
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2MaillingAdd'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2MaiilingAdd2'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2City'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2State'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2ZipCd1'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2ZipCd2'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2Phone'))
    sndAdditionalInsuredVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInsured2Email'))

    fstAdditionalInterestKeys = [
        '1st Additional Interest Interest Type',
        '1st Additional Interest First',
        '1st Additional Interest M.I.',
        '1st Additional Interest Last (or Company)',
        '1st Additional Interest International Address',
        '1st Additional Interest Mailing Address',
        '1st Additional Interest Mailing Address Line 2',
        '1st Additional Interest Mailing Address City',
        '1st Additional Interest Mailing Address State',
        '1st Additional Interest Mailing Address Zip',
        '1st Additional Interest Mailing Address Zip2',
        '1st Additional Interest Primary Phone',
        '1st Additional Interest E-mail'
    ]
    fstAdditionalInterestVals = []
    fstAdditionalInterestTag = soup.find(id="ContentPlaceHolderBody_ApplicantMailingBody_Table11")
    fstAdditionalInterestVals.append(getElementbySpan(fstAdditionalInterestTag, "Interest Type"))
    # fstAdditionalInterestVals.append(getEle(soup,'select','ContentPlaceHolderBody_ApplicantMailingBody_ddlAdd1InterestType'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1Name'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1MiddleName'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1LastName'))
    fstAdditionalInterestInternationalAddress = "", fstAdditionalInterestVals.append("")
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1Mailing'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1Mailing2'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1City'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1State'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1ZipCd1'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1ZipCd2'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1Phone'))
    fstAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest1Email'))

    sndAdditionalInterestKeys = [
        '2nd Additional Interest Interest Type','2nd Additional Interest First',
        '2nd Additional Interest M.I.','2nd Additional Interest Last (or Company)',
        '2nd Additional Interest International Address','2nd Additional Interest Mailing Address',
        '2nd Additional Interest Mailing Address Line 2','2nd Additional Interest Mailing Address City',
        '2nd Additional Interest Mailing Address State','2nd Additional Interest Mailing Address Zip',
        '2nd Additional Interest Mailing Address Zip2','2nd Additional Interest Primary Phone',
        '2nd Additional Interest E-mail',
    ]
    sndAdditionalInterestVals = []
    sndAdditionalInterestTag = soup.find(id="ContentPlaceHolderBody_ApplicantMailingBody_Table13")
    sndAdditionalInterestVals.append(getElementbySpan(sndAdditionalInterestTag, "Interest Type"))
    # sndAdditionalInterestVals.append(getEle(soup,'select','ContentPlaceHolderBody_ApplicantMailingBody_ddlAdd2InterestType'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInteresr2Name'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInteresr2MiddleName'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInteresr2LastName'))
    sndAdditionalInterestInternationalAddress = "", sndAdditionalInterestVals.append("")
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest2Mailing'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest2Mailing2'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest2City'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddinterest2State'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest2ZipCd1'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest2ZipCd2'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest2Phone'))
    sndAdditionalInterestVals.append(getEle(soup, 'input', 'ContentPlaceHolderBody_ApplicantMailingBody_txtAddInterest2Email'))
    """End Applicant"""
    totalKeys = applicantKeys + coApplicantKeys + propertyAddressKeys + fstMortageeKeys + sndMortageeKeys + fstAdditionalInsuredKeys + sndAdditionalInsuredKeys + fstAdditionalInterestKeys + sndAdditionalInterestKeys
    totalVals = applicantVals + coApplicantVals + propertyAddressVals + fstMortageeVals + sndMortageeVals + fstAdditionalInsuredVals + sndAdditionalInsuredVals + fstAdditionalInterestVals + sndAdditionalInterestVals
    return {"key":totalKeys, "val":totalVals}

def getGeneral(session,searchKey):
    global product, policyUrl, header
    global RadScriptManager1_TSM, RadStyleSheetManager1_TSSM, __VIEWSTATEFIELDCOUNT, __VIEWSTATEGENERATOR
    session = gotoActionItemsPage(session, searchKey)

    if product == "HO3FL" or product == "DP3FL":index = 3
    else:index = 2
    postData = {
        'ctl00$RadScriptManager1': 'ctl00$ctl00$ContentPlaceHolderBody$trsPolicyPanel|ctl00$ContentPlaceHolderBody$trsPolicy',
        'RadScriptManager1_TSM': RadScriptManager1_TSM, 'RadStyleSheetManager1_TSSM': RadStyleSheetManager1_TSSM,
        '__EVENTTARGET': 'ctl00$ContentPlaceHolderBody$trsPolicy',
        '__EVENTARGUMENT': '{"type":0,"index":"' + str(index) + '"}',
        '__VIEWSTATEFIELDCOUNT': __VIEWSTATEFIELDCOUNT, '__VIEWSTATEGENERATOR': __VIEWSTATEGENERATOR,
        'ContentPlaceHolderBody_trsPolicy_ClientState': '{"selectedIndexes":["' + str(index) + '"],"logEntries":[],"scrollState":{}}',
        'ContentPlaceHolderBody_rmpPolicy_ClientState': '{"selectedIndex":' + str(index) + ',"changeLog":[]}',
    }
    for i in range(int(__VIEWSTATEFIELDCOUNT)):
        if i == 0:postData.update({'__VIEWSTATE': __VIEWSTATES.get("__VIEWSTATE")})
        else:postData.update({'__VIEWSTATE' + str(i): __VIEWSTATES.get("__VIEWSTATE" + str(i))})

    res = session.post(policyUrl, data=postData, headers=header)
    soup = BeautifulSoup(res.text, "html.parser")

    labelEles = soup.find_all('label')
    general = {}
    for label in labelEles:
        val = getElementbyLabel2(soup, label.text)
        if val == -1:
            general[getCleanLabel(label.text)] = ""
        else:
            general[getCleanLabel(label.text)] = val
    return general

def getBuilding(session, searchKey):
    global product, policyUrl, header
    global RadScriptManager1_TSM, RadStyleSheetManager1_TSSM, __VIEWSTATEFIELDCOUNT, __VIEWSTATEGENERATOR

    session = gotoActionItemsPage(session, searchKey)
    postData = getPostData()

    if product == "HO3FL" or product == "DP3FL":
        postData['__EVENTARGUMENT'] = '{"type":0,"index":"1"}'
        postData['ContentPlaceHolderBody_trsPolicy_ClientState'] = '{"selectedIndexes":["1"],"logEntries":[],"scrollState":{}}'
        postData['ContentPlaceHolderBody_rmpPolicy_ClientState'] = '{"selectedIndex":1,"changeLog":[]}'
        res = session.post(policyUrl, data=postData, headers=header)
        soup = BeautifulSoup(res.text, "html.parser")
        DwellingSpans = [
            'Construction Year',
            'Construction Type',
            'Foundation Type',
            'Roof Type',
            'BCEG',
            'Protection Class',
            'Is feet to hydrant less than 1000ft?',
            'Responding Fire Station',
            'Number Of Stories',
            'Number of Full Bathrooms',
            'Number of Half Bathrooms',
            'Total Finished Square Footage',
            'Percent of Lowest Level Finished',
            'Number of Fireplaces',
            'Quality Grade',
            'Exterior Wall Finish',
            'Exterior Wall Construction',
        ]
        DwellingVals = []
        buildingOtherSpans = [
            'Residence Type',
            'Usage Type',
            'Familiy Units in Building',
            'Occupancy',
            'Electrical',
            'Plumbing',
            'Water Heater',
            'Heating',
            'Roofing',
            'Replacement Cost',
            'Estimated Market Value',
            'Actual Insurance Score Band',

            'Fire Protective Devices',
            'Burglar Alarm',
            'Water Protective Devices',
            'Secured Community',
            'Accredited Builder',
            'Supplemental Heating Device'
        ]
        buildingOtherVals = []
        for label in DwellingSpans:
            DwellingVals.append(getElementbySpan2(soup, label))
        for label in buildingOtherSpans:
            buildingOtherVals.append(getElementbySpan(soup, label))
        buildingVals = DwellingVals + buildingOtherVals
        buildingLabel = [
            'Located in windpool eligible area',
            'Distance to Coast',
            'Wind Speed',
            'Terrain',
            'Located in High Velocity Hurricane Zone?',
            'Exclude Wind',
            'WBDR',
            'Is property Located on a barrier island?',
            'Is property located on a seawall?',
            'Florida Building Code',
            'Predominant Roof Covering',
            'Roof Deck Attachment',
            'Roof to Wall Attachment',
            'Roof Geometry',
            'Gable End Bracing',
            'Secondary Water Resistance (SWR)',
            'Opening Protection',
            'Inspection Date',
            'Inspection Company'
        ]
        for label in buildingLabel:buildingVals.append(getElementbyLabel(soup, label))
        buildingKeys = DwellingSpans + buildingOtherSpans + buildingLabel

        """get Coverages in HO3FL, DP3FL"""
        session = gotoActionItemsPage(session,searchKey)
        postData['__EVENTARGUMENT'] = '{"type":0,"index":"2"}'
        postData['ContentPlaceHolderBody_trsPolicy_ClientState'] = '{"selectedIndexes":["2"],"logEntries":[],"scrollState":{}}'
        postData['ContentPlaceHolderBody_rmpPolicy_ClientState'] = '{"selectedIndex":2,"changeLog":[]}'
        res = session.post(policyUrl, data=postData, headers=header)
        soup = BeautifulSoup(res.text, "html.parser")
        if product == "HO3FL":
            coverageKeys = [
                'Dwelling - A',
                'Other Structures - B',
                'Personal Property - C',
                'Loss of Use - D',
                'Personal Liability - E',
                'Medical Payments - F',
                'All Other Perils',
                'Hurricane',
                'Personal Property Replacement Cost',
                'Specified Additional Amount of Insurance - Coverage A',
                'Water Back Up and Sump Overflow',
                'Loss Assessment',
                'Ordinance or Law',
                'Special Personal Property',
                'Personal Injury',
                'Water Damage Exclusion',
                'Direct Repair',
                'Limited Fungi, Mold, Wet or Dry Rot, or Bacteria',
                'Identity Fraud Expense',
                'Hurricane Screened Enclosure',
                'Golf Cart',
                '# of Carts',
                'Animal Liability',
                'Sinkhole',
                'Limited Water Damage'
            ]
        else:
            coverageKeys = [
            'Dwelling - A',
            'Actual Cash Value Loss Settlement',
            'Other Structures - B',
            'Personal Property - C',
            'Fair Rental Value - D',
            'Additional Living Expense - E',
            'Personal Liability - L',
            'Medical Payments - M',
            'All Other Perils',
            'Hurricane',
            'Personal Property Replacement Cost - Coverage C',
            'Loss Assessment',
            'Ordinance or Law',
            'Permitted Incidental Occupancy',
            'Direct Repair',
            'Water Damage Exclusion',
            'Limited Fungi, Mold, Wet or Dry Rot, or Bacteria Increased Coverage',
            'Theft Coverage',
            'Wind or Hail - Screened Enclosures and Carports',
            'Sinkhole',
            'Limited Water Damage'
        ]
        coverageVals = []
        for label in coverageKeys:
            coverageVals.append(getElementbyLabel(soup, label))

        totalKeys = buildingKeys + coverageKeys
        totalVals = buildingVals + coverageVals
        return {"key":totalKeys, "val":totalVals}
    ####"""get Policy in HO6FL and HO3TX"""
    elif product == "HO6FL" or product == "HO3TX":
        postData['__EVENTARGUMENT'] = '{"type":0,"index":"1"}'
        postData['ContentPlaceHolderBody_trsPolicy_ClientState'] = '{"selectedIndexes":["1"],"logEntries":[],"scrollState":{}}'
        postData['ContentPlaceHolderBody_rmpPolicy_ClientState'] = '{"selectedIndex":1,"changeLog":[]}'
        """get Policy data"""
        res = session.post(policyUrl, data=postData, headers=header)
        soup = BeautifulSoup(res.text, "html.parser")
        labelEles = soup.find_all('label')
        polKeys = []
        polVals = []
        for label in labelEles:
            val = getElementbyLabel(soup, label.text)
            if val == -1 or label.text.strip() == '': continue
            polKeys.append(getCleanLabel(label.text))
            polVals.append(val)
        return {"key":polKeys, "val":polVals}
    else:return {"key":[],"val":[]}

"""
    added on 01/03/2019
"""
def getLossHistory(session, searchKey):
    global product, policyUrl, header
    global RadScriptManager1_TSM, RadStyleSheetManager1_TSSM, __VIEWSTATEFIELDCOUNT, __VIEWSTATEGENERATOR

    session = gotoActionItemsPage(session, searchKey)
    postData = getPostData()

    if product == "HO3FL" or product == "DP3FL":
        postData['__EVENTARGUMENT'] = '{"type":0,"index":"4"}'
        postData['ContentPlaceHolderBody_trsPolicy_ClientState'] = '{"selectedIndexes":["4"],"logEntries":[],"scrollState":{}}'
        postData['ContentPlaceHolderBody_rmpPolicy_ClientState'] = '{"selectedIndex":4,"changeLog":[]}'
    else:
        postData['__EVENTARGUMENT'] = '{"type":0,"index":"3"}'
        postData['ContentPlaceHolderBody_trsPolicy_ClientState'] = '{"selectedIndexes":["3"],"logEntries":[],"scrollState":{}}'
        postData['ContentPlaceHolderBody_rmpPolicy_ClientState'] = '{"selectedIndex":5,"changeLog":[]}'

    res = session.post(policyUrl, data=postData, headers=header)
    soup = BeautifulSoup(res.text, "html.parser")
    data = []
    div = soup.find('div', id='ContentPlaceHolderBody_LossHistoryBody_pnlLossHistory')
    table = div.find_all('table', id='ContentPlaceHolderBody_LossHistoryBody_TableClaimHistory')
    if table:
        table1 = table.find('table', id='ContentPlaceHolderBody_LossHistoryBody_RadGridLossHistory_ctl00')
        if table1:
            trs = table1.find_all('tr')
            if len(trs) > 1:
                for tr in trs:
                    tds = tr.find_all('td')
                    if len(tds) > 5:
                        ClaimNumber = tds[0].text
                        LossDate = tds[1].text
                        LossLocation = tds[2].text
                        LossType = tds[3].text
                        LossAmount = tds[4].text
                        PriorClaimStatus = tds[5].text
                        dt = {
                            "Claim Number" : ClaimNumber,
                            "Loss Date" : LossDate,
                            "Loss Location" : LossLocation,
                            "Loss Type" : LossType,
                            "Loss Amount" : LossAmount,
                            "Prior Claim Status" : PriorClaimStatus
                        }
                        data.append(dt)
    return {"key":["LossHistory"],"val":data}

def Scrapping_Unit(searchKey):
    session = login()
    session = getPolicyBilling(session,searchKey)
    ApplicantData = getApplicant(session)
    """get General"""
    general = getGeneral(session,searchKey)
    """get Building in HO3FL, DP3FL"""
    BuildingData = getBuilding(session,searchKey)
    """ get LossHistory """
    LossHistory = getLossHistory(session,searchKey)

    """export data to csv"""
    csvfile1 = "HO3FL.csv"
    csvfile2 = "HO6FL.csv"
    csvfile3 = "DP3FL.csv"
    csvfile4 = "HO6FL.csv"
    csvhead = ["Policy #"] + policyKeys + ApplicantData['key'] + ['general'] + BuildingData['key'] + LossHistory['key']
    csvdata = [searchKey] + policyVals + ApplicantData['val'] + [general] + BuildingData['val'] + LossHistory['val']
    if product == "HO3FL":
        outputCSV(csvfile1, csvhead, csvdata)
    elif product == "DP3FL":
        outputCSV(csvfile3,csvhead, csvdata)
    elif product == "HO6FL":
        outputCSV(csvfile2, csvhead, csvdata)
    elif product == "HO3TX":
        outputCSV(csvfile4, csvhead, csvdata)

if __name__ == "__main__":
    # excel_download()
    file = pd.read_excel("policykey.xls",sheet_name='RadGridExport')
    searchKeys = file['Policy #']._ndarray_values
    for key in searchKeys:
        Scrapping_Unit(key)
        print(key)


