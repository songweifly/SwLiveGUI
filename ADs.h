#include <Iads.h>  
#include <Adshlp.h>  
#include <activeds.h>  
#include <adserr.h>  
#include <comutil.h>  
#pragma comment(lib,"Activeds.lib")  
#pragma comment(lib,"adsiid.lib")  
#pragma comment(lib,"comsuppw.lib")  
//  Define UNICODE.  
//  Define version 5 for Windows 2000.  
#define _WIN32_WINNT 0x0500  
#include <sddl.h>  
  
HRESULT FindUsers(IDirectorySearch *pContainerToSearch,  //  IDirectorySearch pointer to the container to search.  
                  LPOLESTR szFilter, //  Filter to find specific users.  
                  //  NULL returns all user objects.  
                  LPOLESTR *pszPropertiesToReturn, //  Properties to return for user objects found.  
                  //  NULL returns all set properties.  
                  BOOL bIsVerbose //  TRUE indicates that display all properties for the found objects.  
                  //  FALSE indicates that only the RDN.  
                  ); 


HRESULT CreateUserFromADs(LPCWSTR pwszContainerDN,
                          LPCWSTR pwszName, 
                          LPCWSTR pwszSAMAccountName, 
                          LPCWSTR pwszInitialPassword);



LPWSTR ConvertCharToLPWSTR(const char * szString);
CString ConvertCharToSTR(LPCWSTR strWide);