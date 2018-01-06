  
#include "stdafx.h" 
#include "ADs.h"
  
 
  
HRESULT FindUsers(IDirectorySearch *pContainerToSearch,  //  IDirectorySearch pointer to the container to search.  
                  LPOLESTR szFilter, //  Filter for finding specific users.  
                  //  NULL returns all user objects.  
                  LPOLESTR *pszPropertiesToReturn, //  Properties to return for user objects found.  
                  //  NULL returns all set properties.  
                  BOOL bIsVerbose    //  TRUE indicates that all properties for the found objects are displayed.  
                  //  FALSE indicates only the RDN.  
                  )  
{  
    if (!pContainerToSearch)  
        return E_POINTER;  
    DWORD dwLength = MAX_PATH*2;  
    // Create search filter.  
    LPOLESTR pszSearchFilter = new OLECHAR[dwLength];  
  
    //  Add the filter.  
    swprintf_s(pszSearchFilter, dwLength, L"(&(objectClass=user)(objectCategory=person)%s)",szFilter); 
  
    //  Specify subtree search.  
    ADS_SEARCHPREF_INFO SearchPrefs;  
    SearchPrefs.dwSearchPref = ADS_SEARCHPREF_SEARCH_SCOPE;  
    SearchPrefs.vValue.dwType = ADSTYPE_INTEGER;  
    SearchPrefs.vValue.Integer = ADS_SCOPE_SUBTREE;  
    DWORD dwNumPrefs = 1;  
  
    //  COL for iterations.  
    LPOLESTR pszColumn = NULL;    
    ADS_SEARCH_COLUMN col;  
    HRESULT hr = S_OK;  
  
  
    //  Search handle.  
    ADS_SEARCH_HANDLE hSearch = NULL;  

	//File
	CStdioFile file;
	if(!file.Open("ADs.txt", CFile::modeCreate|CFile::modeWrite|CFile::typeText)) 
	{
		AfxMessageBox("打开表配置文件失败！");
		return hr;
	}
  
    //  Set search preference.  
    hr = pContainerToSearch->SetSearchPreference(&SearchPrefs, dwNumPrefs);  
    if (FAILED(hr))  
        return hr;  
  
    LPOLESTR pszBool = NULL;  
    DWORD dwBool = 0;  
    PSID pObjectSID = NULL;  
    LPOLESTR szSID = NULL;  
    LPOLESTR szDSGUID = new WCHAR [39];  
    LPGUID pObjectGUID = NULL;  
    FILETIME filetime;  
    SYSTEMTIME systemtime;  
    DATE date;  
    VARIANT varDate;  
    LARGE_INTEGER liValue;  
    LPOLESTR *pszPropertyList = NULL;  
    //LPOLESTR pszNonVerboseList[] = {L"name",L"distinguishedName"};  
	LPOLESTR pszNonVerboseList[] = {L"displayName",L"userPrincipalName",L"userAccountControl"};  
  
    LPOLESTR szName = new OLECHAR[MAX_PATH];  
    LPOLESTR szDN = new OLECHAR[MAX_PATH];  
	 LPOLESTR szPwd = new OLECHAR[MAX_PATH]; 
  
    VariantInit(&varDate);  
  
    int iCount = 0;  
    DWORD x = 0L;  
  
  
  
    if (!bIsVerbose)  
    {  
        //  Return non-verbose list properties only.  
        hr = pContainerToSearch->ExecuteSearch(pszSearchFilter,  
            pszNonVerboseList,  
            sizeof(pszNonVerboseList)/sizeof(LPOLESTR),  
            &hSearch  
            );  
    }  
    else  
    {  
        if (!pszPropertiesToReturn)  
        {  
            //  Return all properties.  
            hr = pContainerToSearch->ExecuteSearch(pszSearchFilter,  
                NULL,  
                (DWORD)-1,  
                &hSearch  
                );  
        }  
        else  
        {  
            //  Specified subset.  
            pszPropertyList = pszPropertiesToReturn;  
            //  Return specified properties.  
            hr = pContainerToSearch->ExecuteSearch(pszSearchFilter,  
                pszPropertyList,  
                sizeof(pszPropertyList)/sizeof(LPOLESTR),  
                &hSearch  
                );  
        }  
    }  
    if (SUCCEEDED(hr))  
    {    
        //  Call IDirectorySearch::GetNextRow() to retrieve the next data row.  
        hr = pContainerToSearch->GetFirstRow(hSearch);  
        if (SUCCEEDED(hr))  
        {  
            while(hr != S_ADS_NOMORE_ROWS)  
            {  
                //  Keep track of count.  
                iCount++;  
                //  Loop through the array of passed column names,  
                //  print the data for each column.  
  
                while(pContainerToSearch->GetNextColumnName(hSearch, &pszColumn) != S_ADS_NOMORE_COLUMNS)  
                {  
                    hr = pContainerToSearch->GetColumn(hSearch, pszColumn, &col);  
                    if (SUCCEEDED(hr))  
                    {  
                        //  Print the data for the column and free the column.  
                        if(bIsVerbose)  
                        {  
                            //  Get the data for this column.  
                            wprintf(L"%s\n",col.pszAttrName);  
                            switch (col.dwADsType)  
                            {  
                            case ADSTYPE_DN_STRING:  
                                for (x = 0; x< col.dwNumValues; x++)  
                                {  
                                    wprintf(L"  %s\r\n",col.pADsValues[x].DNString);  
                                }  
                                break;  
                            case ADSTYPE_CASE_EXACT_STRING:      
                            case ADSTYPE_CASE_IGNORE_STRING:      
                            case ADSTYPE_PRINTABLE_STRING:      
                            case ADSTYPE_NUMERIC_STRING:        
                            case ADSTYPE_TYPEDNAME:          
                            case ADSTYPE_FAXNUMBER:          
                            case ADSTYPE_PATH:            
                                for (x = 0; x< col.dwNumValues; x++)  
                                {  
                                    wprintf(L"  %s\r\n",col.pADsValues[x].CaseIgnoreString);  
                                }  
                                break;  
                            case ADSTYPE_BOOLEAN:  
                                for (x = 0; x< col.dwNumValues; x++)  
                                {  
                                    dwBool = col.pADsValues[x].Boolean;  
                                    pszBool = dwBool ? L"TRUE" : L"FALSE";  
                                    wprintf(L"  %s\r\n",pszBool);  
                                }  
                                break;  
                            case ADSTYPE_INTEGER:  
                                for (x = 0; x< col.dwNumValues; x++)  
                                {  
                                    wprintf(L"  %d\r\n",col.pADsValues[x].Integer);  
                                }  
                                break;  
                            case ADSTYPE_OCTET_STRING:  
                                if (_wcsicmp(col.pszAttrName,L"objectSID") == 0)  
                                {  
                                    for (x = 0; x< col.dwNumValues; x++)  
                                    {  
                                       pObjectSID = (PSID)(col.pADsValues[x].OctetString.lpValue);
                                       CString strTmp = ConvertCharToSTR(szSID); 
                                       char* pstr = strTmp.GetBuffer(strTmp.GetLength());
                                       ConvertSidToStringSid(pObjectSID, &pstr);	
									   strTmp.ReleaseBuffer();
									   szSID = ConvertCharToLPWSTR(strTmp);
                                        wprintf(L"  %s\r\n",szSID);  
                                        LocalFree(szSID);  
                                    }  
                                }  
                                else if ((_wcsicmp(col.pszAttrName,L"objectGUID") == 0))  
                                {  
                                    for (x = 0; x< col.dwNumValues; x++)  
                                    {  
                                        //  Cast to LPGUID.  
                                        pObjectGUID = (LPGUID)(col.pADsValues[x].OctetString.lpValue);  
                                        //  Convert GUID to string.  
                                        ::StringFromGUID2(*pObjectGUID, szDSGUID, 39);   
                                        //  Print the GUID.  
                                        wprintf(L"  %s\r\n",szDSGUID);  
                                    }  
                                }  
                                else  
                                    wprintf(L"  Value of type Octet String. No Conversion.");  
                                break;  
                            case ADSTYPE_UTC_TIME:  
                                for (x = 0; x< col.dwNumValues; x++)  
                                {  
                                    systemtime = col.pADsValues[x].UTCTime;  
                                    if (SystemTimeToVariantTime(&systemtime,  
                                        &date) != 0)   
                                    {  
                                        //  Pack in variant.vt.  
                                        varDate.vt = VT_DATE;  
                                        varDate.date = date;  
                                        VariantChangeType(&varDate,&varDate,VARIANT_NOVALUEPROP,VT_BSTR);  
                                        wprintf(L"  %s\r\n",varDate.bstrVal);  
                                        VariantClear(&varDate);  
                                    }  
                                    else  
                                        wprintf(L"  Could not convert UTC-Time.\n",pszColumn);  
                                }  
                                break;  
                            case ADSTYPE_LARGE_INTEGER:  
                                for (x = 0; x< col.dwNumValues; x++)  
                                {  
                                    liValue = col.pADsValues[x].LargeInteger;  
                                    filetime.dwLowDateTime = liValue.LowPart;  
                                    filetime.dwHighDateTime = liValue.HighPart;  
                                    if((filetime.dwHighDateTime==0) && (filetime.dwLowDateTime==0))  
                                    {  
                                        wprintf(L"  No value set.\n");  
                                    }  
                                    else  
                                    {  
                                        //  Verify properties of type LargeInteger that represent time.  
                                        //  If TRUE, then convert to variant time.  
                                        if ((0==wcscmp(L"accountExpires", col.pszAttrName))|  
                                            (0==wcscmp(L"badPasswordTime", col.pszAttrName))||  
                                            (0==wcscmp(L"lastLogon", col.pszAttrName))||  
                                            (0==wcscmp(L"lastLogoff", col.pszAttrName))||  
                                            (0==wcscmp(L"lockoutTime", col.pszAttrName))||  
                                            (0==wcscmp(L"pwdLastSet", col.pszAttrName))  
                                            )  
                                        {  
                                            //  Handle special case for Never Expires where low part is -1.  
                                            if (filetime.dwLowDateTime==-1)  
                                            {  
                                                wprintf(L"  Never Expires.\n");  
                                            }  
                                            else  
                                            {  
                                                if (FileTimeToLocalFileTime(&filetime, &filetime) != 0)   
                                                {  
                                                    if (FileTimeToSystemTime(&filetime,  
                                                        &systemtime) != 0)  
                                                    {  
                                                        if (SystemTimeToVariantTime(&systemtime,  
                                                            &date) != 0)   
                                                        {  
                                                            //  Pack in variant.vt.  
                                                            varDate.vt = VT_DATE;  
                                                            varDate.date = date;  
                                                            VariantChangeType(&varDate,&varDate,VARIANT_NOVALUEPROP,VT_BSTR);  
                                                            wprintf(L"  %s\r\n",varDate.bstrVal);  
                                                            VariantClear(&varDate);  
                                                        }  
                                                        else  
                                                        {  
                                                            wprintf(L"  FileTimeToVariantTime failed\n");  
                                                        }  
                                                    }  
                                                    else  
                                                    {  
                                                        wprintf(L"  FileTimeToSystemTime failed\n");  
                                                    }  
  
                                                }  
                                                else  
                                                {  
                                                    wprintf(L"  FileTimeToLocalFileTime failed\n");  
                                                }  
                                            }  
                                        }  
                                        else  
                                        {  
                                            //  Print the LargeInteger.  
                                            wprintf(L"  high: %d low: %d\r\n",filetime.dwHighDateTime, filetime.dwLowDateTime);  
                                        }  
                                    }  
                                }  
                                break;  
                            case ADSTYPE_NT_SECURITY_DESCRIPTOR:  
                                for (x = 0; x< col.dwNumValues; x++)  
                                {  
                                    wprintf(L"  Security descriptor.\n");  
                                }  
                                break;  
                            default:  
                                wprintf(L"Unknown type %d.\n",col.dwADsType);  
                            }  
                        }  
                        else  
                        {  
                            //  Verbose handles only the two single-valued attributes: cn and ldapdisplayname,  
                            //  so this is a special case.  
                            //if (0==wcscmp(L"name", pszColumn)) 
							if (0==wcscmp(L"displayName", pszColumn))  
                            {  
                                wcscpy(szName,col.pADsValues->CaseIgnoreString);  
                            }  
                            //if (0==wcscmp(L"distinguishedName", pszColumn))  
							if (0==wcscmp(L"userPrincipalName", pszColumn))  
                            {  
                                wcscpy(szDN,col.pADsValues->CaseIgnoreString);  
                            }  
							if (0==wcscmp(L"userAccountControl", pszColumn))  
                            {  
                               // wcscpy(szPwd,col.pADsValues->Integer); 
								if( col.pADsValues->Integer == 66050)
									wcscpy(szPwd, L"禁用");
								else
									wcscpy(szPwd, L"正常");
                            } 
                        }  
                        pContainerToSearch->FreeColumn(&col);  
                    }  
                    FreeADsMem(pszColumn);  
                }  
                if (!bIsVerbose)
				{
					//wprintf(L"%s\n  DN: %s\n\n",szName,szDN);
					CString strName(szName);
					CString strDN(szDN);
					CString strPwd(szPwd);
					CString strLine;
					strLine.Format("%s|%s|%s\n", strName,strDN,strPwd);
					file.WriteString(strLine); 
				}
                //  Get the next row.  
                hr = pContainerToSearch->GetNextRow(hSearch);  
            }  
  
        }  
        //  Close the search handle to cleanup.  
        pContainerToSearch->CloseSearchHandle(hSearch);  
    }   
    if (SUCCEEDED(hr) && 0==iCount)  
        hr = S_FALSE;  
  
	file.Close();
    delete [] szName;  
    delete [] szDN;  
    delete [] szDSGUID;  
    delete [] pszSearchFilter;  
    return hr;  
}  

LPWSTR  ConvertCharToLPWSTR(const char * szString)
{
	int dwLen = strlen(szString) + 1;
	int nwLen = MultiByteToWideChar(CP_ACP, 0, szString, dwLen, NULL, 0);//算出合适的长度
	LPWSTR lpszPath = new WCHAR[dwLen];
	MultiByteToWideChar(CP_ACP, 0, szString, dwLen, lpszPath, nwLen);
	return lpszPath;
}
 

CString ConvertCharToSTR(LPCWSTR strWide)
{
  //从宽字符串转换窄字符串 
    //获取转换所需的目标缓存大小
    DWORD dBufSize=WideCharToMultiByte(CP_OEMCP, 0, strWide, -1, NULL,0,NULL, FALSE);

    //分配目标缓存
    char* dBuf = new char[dBufSize];
    memset(dBuf, 0, dBufSize);

    //转换
    int nRet=WideCharToMultiByte(CP_OEMCP, 0, strWide, -1, dBuf, dBufSize, NULL, FALSE);    
    if(nRet<=0)
    {
        printf("转换失败\n");
    }
    else
    {
        printf("转换成功\nAfter Convert: %s\n", dBuf);
    }

	CString str = dBuf;
    delete []dBuf;
	return str;
}



HRESULT CreateUserFromADs(LPCWSTR pwszContainerDN,
                          LPCWSTR pwszName, 
                          LPCWSTR pwszSAMAccountName, 
                          LPCWSTR pwszInitialPassword)
{
    HRESULT hr;

    //  Build the DN of the container.
    CComBSTR sbstrADsPath = "LDAP://";
    sbstrADsPath += pwszContainerDN;

    IADsContainer *pUsers = NULL;


	LPWSTR szUsername = ConvertCharToLPWSTR("administrator"); // user name  
	LPWSTR szPassword = ConvertCharToLPWSTR("Swn123321"); // password  

    // Bind to the container.
    //hr = ADsGetObject(sbstrADsPath, IID_IADsContainer, (LPVOID*)&pUsers);
	hr = ADsOpenObject(L"LDAP://132.147.253.125:389/OU=DEVVDI,DC=DEVVDI,DC=SW,DC=COM",   
                szUsername,  
                szPassword,  
                ADS_SECURE_AUTHENTICATION,  
                IID_IADsContainer,  
                (LPVOID*)&pUsers);  
    if(SUCCEEDED(hr))
    {
        IDispatch *pDisp = NULL;
        
        CComBSTR sbstrName = "CN=";
        sbstrName += pwszName;
        
        // Create the new object in the User folder.
        hr = pUsers->Create(CComBSTR("user"), sbstrName, &pDisp);

        if(SUCCEEDED(hr))
        { 
            IADsUser *padsUser = NULL;

            // Get the IADs interface.
            hr = pDisp->QueryInterface(IID_IADsUser, (void**) &padsUser);

            if(SUCCEEDED(hr))
            { 
                CComBSTR sbstrProp,svar;
                /*
                The sAMAccountName property is required on OS versions 
                prior to Windows Server 2003 family. The Windows Server 2003 
                family will create a sAMAccountName value if one is not 
                specified.
                */
                svar = pwszSAMAccountName;
                sbstrProp = "sAMAccountName";
                hr = padsUser->Put(sbstrProp, svar);

                /*
                Commit the new user to persistent memory. The user does not 
                exist until this is called.
                */
                hr = padsUser->SetInfo();

                /*
                Set the initial password. This must be done after 
                SetInfo is called because the user object must 
                already exist on the server.
                */
                hr = padsUser->SetPassword(CComBSTR(pwszInitialPassword));

                /*
                Set the pwdLastSet property to zero, which forces the 
                user to change the password the next time they log on.
                */
                sbstrProp = "pwdLastSet";
                svar = 0;
                hr = padsUser->Put(sbstrProp, svar);

                /*
                Enable the user account by removing the 
                ADS_UF_ACCOUNTDISABLE flag from the userAccountControl 
                property. Also, remove the ADS_UF_PASSWD_NOTREQD and 
                ADS_UF_DONT_EXPIRE_PASSWD flags from the 
                userAccountControl property.
                */
                svar.Clear();
                sbstrProp = "userAccountControl";
                hr = padsUser->Get(sbstrProp, &svar);
                if(SUCCEEDED(hr))
                {
                    svar = svar.lVal & ~(ADS_UF_ACCOUNTDISABLE | 
                        ADS_UF_PASSWD_NOTREQD | 
                        ADS_UF_DONT_EXPIRE_PASSWD);

                    hr = padsUser->Put(sbstrProp, svar);
                    hr = padsUser->SetInfo();
                }
                
                hr = padsUser->put_AccountDisabled(VARIANT_FALSE);
                hr = padsUser->SetInfo();

                padsUser->Release();
            }

            pDisp->Release();
        }

        pUsers->Release();
    }

    return hr;
}