import requests
from bs4 import BeautifulSoup
import json
import os,sys


username = "eligra1"
password = "VR8469"
policyUrl = "https://policy.velocityrisk.com/Policy/Policy.aspx"
header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36','X-MicrosoftAjax': 'Delta=true', 'X-Requested-With': 'XMLHttpRequest'}
global RadScriptManager1_TSM, RadStyleSheetManager1_TSSM, __VIEWSTATEFIELDCOUNT, __VIEWSTATES,__VIEWSTATEGENERATOR,__EVENTVALIDATION

def getCookie(soup):
    global RadScriptManager1_TSM, RadStyleSheetManager1_TSSM, __VIEWSTATEFIELDCOUNT, __VIEWSTATES,__VIEWSTATEGENERATOR,__EVENTVALIDATION
    try:
        RadScriptManager1_TSM = soup.find('input', {'id': 'RadScriptManager1_TSM'}).get('value')
    except:pass
    try:
        RadStyleSheetManager1_TSSM = soup.find('input', {'id': 'RadStyleSheetManager1_TSSM'}).get('value')
    except:pass
    try:
        __VIEWSTATEFIELDCOUNT = soup.find('input', {'id': '__VIEWSTATEFIELDCOUNT'}).get('value')
    except:pass
    try:
        __VIEWSTATES = {}
    except:pass
    try:
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
    session = requests.session()
    Homeurl = "https://policy.velocityrisk.com/login.aspx?ReturnUrl=%2f"
    res = session.get(Homeurl)
    soup = BeautifulSoup(res.text, 'html.parser')
    try:
        viewState           = soup.find('input', {'id': '__VIEWSTATE'}).get('value')
        viewStateGenerator  = soup.find('input', {'id': '__VIEWSTATEGENERATOR'}).get('value')
        viewEventValidation = soup.find('input', {'id': '__EVENTVALIDATION'}).get('value')
    except:pass

    loginUrl = "https://policy.velocityrisk.com/login.aspx?ReturnUrl=%2f"
    postData = {'AquantLogin$UserName': username,'AquantLogin$Password': password,
                'AquantLogin$LoginButton': 'Log In','__VIEWSTATE': viewState,
                '__VIEWSTATEGENERATOR': viewStateGenerator,'__EVENTVALIDATION' : viewEventValidation}
    res = session.post(loginUrl, data=postData, allow_redirects=False)
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

def PolicyBilling(session, searchKey):
    global policyVals, policyKeys, policyUrl, header
    global RadScriptManager1_TSM, RadStyleSheetManager1_TSSM, __VIEWSTATEFIELDCOUNT, __VIEWSTATES, __VIEWSTATEGENERATOR, __EVENTVALIDATION
    """get policy data & get Hidden fields & Policy History data"""
    ActionUrl = "https://policy.velocityrisk.com/HomeItems/ActionItems.aspx"
    res = session.get(ActionUrl)
    soup = BeautifulSoup(res.text, "html.parser")
    getCookie(soup)

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
    return session

def PolicyDocument(session):
    global header
    global RadScriptManager1_TSM, RadStyleSheetManager1_TSSM, __VIEWSTATEFIELDCOUNT, __VIEWSTATES, __VIEWSTATEGENERATOR, __EVENTVALIDATION

    res = session.get("https://policy.velocityrisk.com/PolicyDocuments.aspx", headers=header)
    soup = BeautifulSoup(res.text, "html.parser")
    pdfNames = []
    try:
        trs = soup.find('table',{'id':'ctl01_RadGridExistingAttachments_ctl00'}).find_all('tbody')[0].find_all('tr')
        for tr in trs:
            tds = tr.find_all('td')
            for td in tds:
                if "pdf" in td.text:
                    pdfNames.append(td.text)
    except:
        pass

    # getCookie(soup)

    postData = {
        'RadScriptManager1_TSM': RadScriptManager1_TSM,
        'RadStyleSheetManager1_TSSM': RadStyleSheetManager1_TSSM,
        '__VIEWSTATEFIELDCOUNT': __VIEWSTATEFIELDCOUNT,
        '__VIEWSTATEGENERATOR': __VIEWSTATEGENERATOR,
        '__EVENTVALIDATION': __EVENTVALIDATION,
        'RadScriptManager1': 'ctl01$ctl01$RadGridExistingAttachmentsPanel|ctl01$RadGridExistingAttachments$ctl00$ctl04$ctl00',
        'ctl01$HiddenFieldMode': 'FileUpload',
        '__EVENTTARGET': 'ctl01$RadGridExistingAttachments$ctl00$ctl04$ctl00',
        'RadAJAXControlID': 'ctl01_RadAjaxManagerPolicyDocs'
    }

    for i in range(int(__VIEWSTATEFIELDCOUNT)):
        if i == 0:postData.update({'__VIEWSTATE': __VIEWSTATES.get("__VIEWSTATE")})
        else:postData.update({'__VIEWSTATE' + str(i): __VIEWSTATES.get("__VIEWSTATE" + str(i))})
    res = session.post("https://policy.velocityrisk.com/PolicyDocuments.aspx", data=postData)

    for name in pdfNames:
        fileName = name
        url = "https://policy.velocityrisk.com/FileServe.aspx?DocumentId=749572&Source=POLICY"
        r = session.get(url)
        with open(fileName, 'wb') as f:
            f.write(r.content)
    return session

def main():
    session = login()
    session = PolicyBilling(session,"VUW-HW-574029")
    session = PolicyDocument(session)


if __name__ == "__main__":
    main()