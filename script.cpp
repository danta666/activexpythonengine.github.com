// Script.Cpp : implementation file
//
// Copyright (c) Microsoft Corporation.  All rights reserved.
//
// This source code is only intended as a supplement to the
// Microsoft Classes Reference and related electronic
// documentation provided with the library.
// See these sources for detailed information regarding the
// Microsoft C++ Libraries products.

#include "StdAfx.H"
#include "TestCon.H"
#include <fstream> 
#include <windows.h>
#include <iostream>

using namespace std;
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#define UNICODE

#endif
/////////////////////////////////////////////////////////////////////////////
// CScript
#define mm L
IMPLEMENT_DYNAMIC( CScript, CCmdTarget )

CScript::CScript( CScriptManager* pManager ) :
   m_pManager( pManager )
{
   m_bPythonScript = false;
   m_bPrintLogEvent = false;
   ASSERT( m_pManager != NULL );
}

CScript::~CScript()
{
   Unload();
}

HRESULT CScript::AddNamedItem( LPCTSTR pszItemName )
{

   HRESULT hResult;

   CT2COLE pszItemNameO( pszItemName );

   hResult = m_pActiveScript->AddNamedItem( pszItemNameO,
	  SCRIPTITEM_ISSOURCE|SCRIPTITEM_ISVISIBLE );
   if( FAILED( hResult ) )
   {
	  return( hResult );
   }

   return( S_OK );
}

BOOL CScript::FindMacro( LPCTSTR pszMacroName )
{
   DISPID dispid;

   ASSERT( pszMacroName != NULL );

   return( m_mapMacros.Lookup( pszMacroName, dispid ) );
}

POSITION CScript::GetFirstMacroPosition()
{
   return( m_mapMacros.GetStartPosition() );
}

CString CScript::GetName()
{
   return( m_strScriptName );
}

CString CScript::GetNextMacroName( POSITION& posMacro )
{
   CString strMacro;
   DISPID dispid;

   m_mapMacros.GetNextAssoc( posMacro, strMacro, dispid );

   return( strMacro );
}

HRESULT CScript::LoadScript( LPCTSTR pszFileName, LPCTSTR pszScriptName )
{
   ITypeInfoPtr pTypeInfo;
   HRESULT hResult;
   CLSID clsid;
   CFile file;
   ULONGLONG nFileSize;
   BSTR bstrScript;
   BOOL tSuccess;
   int iMethod;
   UINT nNames;
   BSTR bstrName;
   CString strName;
   CString strExtension;
   int iChar;

   ENSURE( pszFileName != NULL );
   ENSURE( pszScriptName != NULL );
   ENSURE( m_pActiveScript == NULL );

   iChar = lstrlen( pszFileName )-1;
   while( (iChar >= 0) && (pszFileName[iChar] != _T( '.' )) )
   {
	  iChar--;
   }

   if( (iChar >= 0) && (_tcsicmp( &pszFileName[iChar], _T( ".js" ) ) == 0) )
   {
	  hResult = CLSIDFromProgID( L"JScript", &clsid );
	  if( FAILED( hResult ) )
	  {
		 return( hResult );
	  }
   }
   else if((_tcsicmp( &pszFileName[iChar], _T( ".py" ) ) == 0))
   {//增加对python脚本的支持20180111
	   hResult = CLSIDFromProgID( L"Python", &clsid );
	   if( FAILED( hResult ) )
	   {
		   return( hResult );
	   }
	   m_bPythonScript = true;
   }
   else
   {
	   hResult = CLSIDFromProgID( L"VBScript", &clsid );
	   if( FAILED( hResult ) )
	   {
		   return( hResult );
	   }

   }

   hResult = m_pActiveScript.CreateInstance( clsid, NULL,
	  CLSCTX_INPROC_SERVER );
   if( FAILED( hResult ) )
   {
	  return( hResult );
   }

   const GUID CDECL BASED_CODE _tlid1 =
   { 0x60254CA5, 0x953B, 0x11CF, { 0x8C, 0x96, 0x00, 0xAA, 0x00, 0xB8, 0x70, 0x8C } };

   m_pActiveScriptParse = m_pActiveScript;
   if( m_pActiveScriptParse == NULL )
   {
	  return( E_NOINTERFACE );
   }

   hResult = m_pActiveScript->SetScriptSite( &m_xActiveScriptSite );
   if( FAILED( hResult ) )
   {
	  return( hResult );
   }

   hResult = m_pActiveScriptParse->InitNew();
   if( FAILED( hResult ) )
   {
	  return( hResult );
   }

   CFileException error;
   tSuccess = file.Open( pszFileName, CFile::modeRead|CFile::shareDenyWrite,
	  &error );
   if( !tSuccess )
   {
	  return( HRESULT_FROM_WIN32( error.m_lOsError ) );
   }

   nFileSize = file.GetLength();
   if( nFileSize > INT_MAX )
   {
      return( E_OUTOFMEMORY );
   }
   nFileSize = file.Read( m_strScriptText.GetBuffer( ULONG( nFileSize ) ), ULONG( nFileSize ) );
   file.Close();
   m_strScriptText.ReleaseBuffer( ULONG( nFileSize ) );
   bstrScript = m_strScriptText.AllocSysString();

   hResult = m_pActiveScriptParse->ParseScriptText( bstrScript, NULL, NULL,
	  NULL, DWORD( this ), 0, SCRIPTTEXT_ISVISIBLE, NULL, NULL );
   SysFreeString( bstrScript );
   if( FAILED( hResult ) )
   {
	  return( hResult );
   }

   hResult = m_pManager->AddNamedItems( m_pActiveScript );
   if( FAILED( hResult ) )
   {
	  return( hResult );
   }

   hResult = m_pActiveScript->SetScriptState( SCRIPTSTATE_CONNECTED );
   if( FAILED( hResult ) )
   {
	  return( hResult );
   }
   
   hResult = m_pActiveScript->GetScriptDispatch( NULL, &m_pDispatch );
   if( FAILED( hResult ) )
   {
	  return( hResult );
   }

   if (m_bPythonScript)
   {
	   //解析脚本中的函数名
	   char buffer[256];  
	   fstream outFile;  
	   outFile.open(pszFileName,ios::in);  
	   int index;
	   while(!outFile.eof())  
	   {  
		   outFile.getline(buffer,256,'\n');//getline(char *,int,char) 表示该行字符达到256个或遇到换行就结束  
		   string str = buffer;
		   index=str.find("def");
		   if (index != -1 && str.find("__") == -1)
		   {
			   while( (index = str.find(' ',index)) != string::npos)
			   {
				   str.erase(index,1);
			   }
			   // str--"defClearBtnList():"
			   int first  = str.find("def") + 3;
			   int last  = str.find("(") - 3;
			   
			   if (first == -1 || last == -1)
			   {
				   MessageBox(NULL,"请检查python脚本函数是否定义正确","错误提示",MB_OK);
				   return S_FALSE;
			   }

			   string funName = str.substr(first,last);
			   CString strr(funName.c_str());
			   BSTR b = strr.AllocSysString(); 

			   DISPID dispid;
			   m_pDispatch->GetIDsOfNames(IID_NULL,&b,1,GetUserDefaultLCID(),&dispid);
			   m_mapMacros.SetAt(funName.c_str(),dispid);

		   }	
	   }  
	   outFile.close();  
   }
   else //对于VBScript采用如下代码解析
   {

	   //判断是否支持类型库
	   UINT pp= 0;
	   hResult = m_pDispatch->GetTypeInfoCount(&pp);
	   //加载Python脚本不能获取pTypeinfo，未解决20180117
	   hResult = m_pDispatch->GetTypeInfo( 0, GetUserDefaultLCID(), &pTypeInfo );
	   if( FAILED( hResult ) )
	   {
		   return( hResult );
	   }

	   CSmartTypeAttr pTypeAttr( pTypeInfo );
	   CSmartFuncDesc pFuncDesc( pTypeInfo );

	   hResult = pTypeInfo->GetTypeAttr( &pTypeAttr );
	   if( FAILED( hResult ) )
	   {
		   return( hResult );
	   }

	   for( iMethod = 0; iMethod < pTypeAttr->cFuncs; iMethod++ )
	   {
		   hResult = pTypeInfo->GetFuncDesc( iMethod, &pFuncDesc );
		   if( FAILED( hResult ) )
		   {
			   return( hResult );
		   }

		   if( (pFuncDesc->funckind == FUNC_DISPATCH) && (pFuncDesc->invkind ==
			   INVOKE_FUNC) )
		   {
			   bstrName = NULL;
			   hResult = pTypeInfo->GetNames( pFuncDesc->memid, &bstrName, 1,
				   &nNames );
			   if( FAILED( hResult ) )
			   {
				   return( hResult );
			   }
			   ASSERT( nNames == 1 );

			   strName = bstrName;
			   SysFreeString( bstrName );
			   bstrName = NULL;

			   // Macros can't contain underscores, since those denote event handlers
			   if( strName.Find( _T( '_' ) ) == -1 )
			   {
				   m_mapMacros.SetAt( strName, pFuncDesc->memid );
			   }
		   }

		   pFuncDesc.Release();
	   }


   }

   m_strScriptName = pszScriptName;

   return( S_OK );
}

HRESULT CScript::InvokeScriptHelper( LPCTSTR pszEventName, USHORT wFlags, DISPPARAMS* pdpParams,
	VARIANT* pvarResult, EXCEPINFO* pExceptionInfo, UINT* piArgErr)
{
	ITypeInfoPtr pTypeInfo;
	HRESULT hResult;
	UINT nNames;
	_bstr_t bstrName;
	CString strName;
	CString strExtension;
	DISPID dwDispID;
	/*hResult = m_pDispatch->GetTypeInfo( 0, GetUserDefaultLCID(), &pTypeInfo );
	if( FAILED( hResult ) )
	{
		return( hResult );
	}

	CSmartTypeAttr pTypeAttr( pTypeInfo );
	CSmartFuncDesc pFuncDesc( pTypeInfo );

	hResult = pTypeInfo->GetTypeAttr( &pTypeAttr );
	if( FAILED( hResult ) )
	{
		return( hResult );
	}*/
	BOOL tFound;
	tFound = m_mapMacros.Lookup( pszEventName, dwDispID );
	if( !tFound )
	{
		return( DISP_E_MEMBERNOTFOUND );
	}
	/*bstrName = pszEventName;
	nNames = 1;*/
	//hResult = pTypeInfo->GetIDsOfNames(bstrName.GetAddress(), nNames, &dwDispID);
	//if( FAILED( hResult ) )
	//{
	//return( hResult );
	//}
	//SysFreeString( bstrName );
	//bstrName = NULL;

	// make the call
	SCODE sc = m_pDispatch->Invoke(dwDispID, IID_NULL, 0, wFlags,
		pdpParams, pvarResult, pExceptionInfo, piArgErr);
	
	return( sc );
}

HRESULT CScript::RunMacro( LPCTSTR pszMacroName )
{
   DISPID dispid;
   COleDispatchDriver driver;
   BOOL tFound;

   ASSERT( pszMacroName != NULL );

   tFound = m_mapMacros.Lookup( pszMacroName, dispid );
   if( !tFound )
   {
	  return( DISP_E_MEMBERNOTFOUND );
   }

   driver.AttachDispatch( m_pDispatch, FALSE );
   try
   {
	  driver.InvokeHelper( dispid, DISPATCH_METHOD, VT_EMPTY, NULL, VTS_NONE );
   }
   catch( COleDispatchException* pException )
   {
	  CString strMessage;

	  AfxFormatString1( strMessage, IDS_DISPATCHEXCEPTION,
		 pException->m_strDescription );
	  AfxMessageBox( strMessage );
	  pException->Delete();
   }
   catch( COleException* pException )
   {
	  pException->Delete();
   }

   return( S_OK );
}

void CScript::Unload()
{
   if( m_pDispatch != NULL )
   {
	  m_pDispatch.Release();
   }
   if( m_pActiveScript != NULL )
   {
	  m_pActiveScript->SetScriptState( SCRIPTSTATE_DISCONNECTED );
	  m_pActiveScript->Close();
	  m_pActiveScript.Release();
   }
   if( m_pActiveScriptParse != NULL )
   {
	  m_pActiveScriptParse.Release();
   }
}


BEGIN_INTERFACE_MAP( CScript, CCmdTarget )
   INTERFACE_PART( CScript, IID_IActiveScriptSite, ActiveScriptSite )
   INTERFACE_PART( CScript, IID_IActiveScriptSiteWindow, ActiveScriptSiteWindow )
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(CScript, CCmdTarget)
	//{{AFX_MSG_MAP(CScript)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CScript message handlers


///////////////////////////////////////////////////////////////////////////////
// IActiveScriptSite
///////////////////////////////////////////////////////////////////////////////

STDMETHODIMP_( ULONG ) CScript::XActiveScriptSite::AddRef()
{
   METHOD_PROLOGUE_( CScript, ActiveScriptSite )

   TRACE( "XActiveScriptSite::AddRef()\n" );

   return( pThis->ExternalAddRef() );
}

STDMETHODIMP_( ULONG ) CScript::XActiveScriptSite::Release()
{
   METHOD_PROLOGUE_( CScript, ActiveScriptSite )

   TRACE( "XActiveScriptSite::Release()\n" );

   return( pThis->ExternalRelease() );
}

STDMETHODIMP CScript::XActiveScriptSite::QueryInterface( REFIID iid,
   void** ppInterface )
{
   METHOD_PROLOGUE_( CScript, ActiveScriptSite )

   return( pThis->ExternalQueryInterface( &iid, ppInterface ) );
}

STDMETHODIMP CScript::XActiveScriptSite::GetDocVersionString(
   BSTR* pbstrVersion )
{
   (void)pbstrVersion;

   METHOD_PROLOGUE( CScript, ActiveScriptSite )

   return( E_NOTIMPL );
}

STDMETHODIMP CScript::XActiveScriptSite::GetItemInfo( LPCOLESTR pszName,
   DWORD dwReturnMask, IUnknown** ppItem, ITypeInfo** ppTypeInfo )
{
   HRESULT hResult;
   METHOD_PROLOGUE( CScript, ActiveScriptSite )

   if( ppItem != NULL )
   {
	  *ppItem = NULL;
   }
   if( ppTypeInfo != NULL )
   {
	  *ppTypeInfo = NULL;
   }

   if( dwReturnMask&SCRIPTINFO_IUNKNOWN )
   {
	  if( ppItem == NULL )
	  {
		 return( E_POINTER );
	  }
   }
   if( dwReturnMask&SCRIPTINFO_ITYPEINFO )
   {
	  if( ppTypeInfo == NULL )
	  {
		 return( E_POINTER );
	  }
   }

   COLE2CT pszNameT( pszName );
   TRACE( "XActiveScriptSite::GetItemInfo( %s )\n", pszNameT );

   if( dwReturnMask&SCRIPTINFO_IUNKNOWN )
   {
	  hResult = pThis->m_pManager->GetItemDispatch( pszNameT, ppItem );
	  if( FAILED( hResult ) )
	  {
		 return( hResult );
	  }
   }
  
   if( dwReturnMask&SCRIPTINFO_ITYPEINFO )
   {
	  pThis->m_pManager->GetItemTypeInfo( pszNameT, ppTypeInfo );
   }

   return( S_OK );
}

STDMETHODIMP CScript::XActiveScriptSite::GetLCID( LCID* plcid )
{
   METHOD_PROLOGUE( CScript, ActiveScriptSite )

   if( plcid == NULL )
   {
	  return( E_POINTER );
   }

   *plcid = GetUserDefaultLCID();

   return( S_OK );
}

STDMETHODIMP CScript::XActiveScriptSite::OnEnterScript()
{
   METHOD_PROLOGUE( CScript, ActiveScriptSite )

   return( S_OK );
}

STDMETHODIMP CScript::XActiveScriptSite::OnLeaveScript()
{
   METHOD_PROLOGUE( CScript, ActiveScriptSite )

   return( S_OK );
}

STDMETHODIMP CScript::XActiveScriptSite::OnScriptError(
   IActiveScriptError* pError )
{

   HRESULT hResult;
   CExcepInfo excep;
   CString strMessage;
   int nResult;
   DWORD dwContext;
   ULONG iLine;
   LONG iChar;
   BSTR bstrSourceLineText;

   METHOD_PROLOGUE( CScript, ActiveScriptSite )

   TRACE( "XActiveScriptSite::OnScriptError()\n" );

   ENSURE( pError != NULL );

   hResult = pError->GetSourcePosition( &dwContext, &iLine, &iChar );
   if( SUCCEEDED( hResult ) )
   {
	  TRACE( "Error at line %u, character %d\n", iLine, iChar );
   }
   bstrSourceLineText = NULL;
   hResult = pError->GetSourceLineText( &bstrSourceLineText );
   if( SUCCEEDED( hResult ) )
   {
	  TRACE( "Source Text: %s\n", COLE2CT( bstrSourceLineText ) );
	  SysFreeString( bstrSourceLineText );
	  bstrSourceLineText = NULL;
   }
   hResult = pError->GetExceptionInfo( &excep );
   if( SUCCEEDED( hResult ) )
   {
	  AfxFormatString2( strMessage, IDS_SCRIPTERRORFORMAT, COLE2CT(
		 excep.bstrSource ), COLE2CT( excep.bstrDescription ) );
	  nResult = AfxMessageBox( strMessage, MB_YESNO );
	  if( nResult == IDYES )
	  {
		 return( S_OK );
	  }
	  else
	  {
		 return( E_FAIL );
	  }
   }

   return( E_FAIL );
}

STDMETHODIMP CScript::XActiveScriptSite::OnScriptTerminate(
   const VARIANT* pvarResult, const EXCEPINFO* pExcepInfo )
{
   (void)pvarResult;
   (void)pExcepInfo;

   METHOD_PROLOGUE( CScript, ActiveScriptSite )

   TRACE( "XActiveScriptSite::OnScriptTerminate()\n" );

   return( S_OK );
}

STDMETHODIMP CScript::XActiveScriptSite::OnStateChange(
   SCRIPTSTATE eState )
{
   METHOD_PROLOGUE( CScript, ActiveScriptSite )

   TRACE( "XActiveScriptSite::OnStateChange()\n" );

   switch( eState )
   {
   case SCRIPTSTATE_UNINITIALIZED:
	  TRACE( "\tSCRIPTSTATE_UNINITIALIZED\n" );
	  break;

   case SCRIPTSTATE_INITIALIZED:
	  TRACE( "\tSCRIPTSTATE_INITIALIZED\n" );
	  break;

   case SCRIPTSTATE_STARTED:
	  TRACE( "\tSCRIPTSTATE_STARTED\n" );
	  break;

   case SCRIPTSTATE_CONNECTED:
	  TRACE( "\tSCRIPTSTATE_CONNECTED\n" );
	  break;

   case SCRIPTSTATE_DISCONNECTED:
	  TRACE( "\tSCRIPTSTATE_DISCONNECTED\n" );
	  break;

   case SCRIPTSTATE_CLOSED:
	  TRACE( "\tSCRIPTSTATE_CLOSED\n" );
	  break;

   default:
	  TRACE( "\tUnknown SCRIPTSTATE value\n" );
	  ASSERT( FALSE );
	  break;
   }

   return( S_OK );
}


///////////////////////////////////////////////////////////////////////////////
// IActiveScriptSiteWindow
///////////////////////////////////////////////////////////////////////////////

STDMETHODIMP_( ULONG ) CScript::XActiveScriptSiteWindow::AddRef()
{
   METHOD_PROLOGUE_( CScript, ActiveScriptSiteWindow )

   TRACE( "XActiveScriptSiteWindow::AddRef()\n" );

   return( pThis->ExternalAddRef() );
}

STDMETHODIMP_( ULONG ) CScript::XActiveScriptSiteWindow::Release()
{
   METHOD_PROLOGUE_( CScript, ActiveScriptSiteWindow )

   TRACE( "XActiveScriptSiteWindow::Release()\n" );

   return( pThis->ExternalRelease() );
}

STDMETHODIMP CScript::XActiveScriptSiteWindow::QueryInterface(
   REFIID iid, void** ppInterface )
{
   METHOD_PROLOGUE_( CScript, ActiveScriptSiteWindow )

   return( pThis->ExternalQueryInterface( &iid, ppInterface ) );
}

STDMETHODIMP CScript::XActiveScriptSiteWindow::EnableModeless(
   BOOL tEnable )
{
   (void)tEnable;

   METHOD_PROLOGUE( CScript, ActiveScriptSiteWindow )

   return( E_NOTIMPL );
}

STDMETHODIMP CScript::XActiveScriptSiteWindow::GetWindow(
   HWND* phWindow )
{
   METHOD_PROLOGUE( CScript, ActiveScriptSiteWindow )

   if( phWindow == NULL )
   {
	  return( E_POINTER );
   }

   *phWindow = AfxGetMainWnd()->GetSafeHwnd();

   return( S_OK );
}
