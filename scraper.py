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


def main():
    session = login()
    

if __name__ == "__main__":
    main()