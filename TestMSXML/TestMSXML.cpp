// TestMSXML.cpp : コンソール アプリケーションのエントリ ポイントを定義します。
//

#include "stdafx.h"
#include <windows.h>

#import <msxml4.dll>
//using namespace MSXML2;

int main(int argc, char *argv[])
{
	if (argc != 3) {
		fprintf(stderr, "ARG_ERROR.\n");
		return -1;
	}

	HRESULT hr = ::CoInitializeEx(NULL, COINIT_MULTITHREADED);

	MSXML2::IXMLDOMDocumentPtr pXMLDoc = NULL;
	MSXML2::IXMLDOMElementPtr pXMLElement = NULL;
	MSXML2::IXMLDOMElementPtr pXMLElement2 = NULL;
	MSXML2::IXMLDOMAttributePtr pXMLAttribute = NULL;
	MSXML2::IXMLDOMNodeListPtr pXMLNodes = NULL;
	MSXML2::IXMLDOMNodePtr pXMLNode = NULL;
	MSXML2::IXMLDOMNodePtr pXMLNode2 = NULL;

	hr = CoCreateInstance(CLSID_DOMDocument, NULL, CLSCTX_INPROC_SERVER, IID_IXMLDOMDocument, (void**)&pXMLDoc);

	pXMLDoc->load(argv[1]);
	if (pXMLDoc == NULL)
	{
		fprintf(stderr, "CANN'T READ XML FILE.\n");
		return -1;
	}
	pXMLElement = pXMLDoc->GetdocumentElement();
	pXMLNodes = pXMLElement->selectNodes(_bstr_t(argv[2]));

	if (pXMLNodes != NULL && pXMLNodes->length)
	{

		pXMLElement2 = pXMLDoc->createElement("elem_test");
		pXMLElement2->text = "element_value";
		pXMLAttribute = pXMLDoc->createAttribute("attr_test1");
		pXMLAttribute->value = "attrval1";
		pXMLElement2->setAttributeNode(pXMLAttribute);
		pXMLElement2->setAttribute("attr_test2", "attrval2");

		pXMLAttribute.Release();
		pXMLAttribute = NULL;

		long nodeCnt = pXMLNodes->length;
		for (long i = 0; i < nodeCnt; ++i) {
			pXMLNode = pXMLNodes->item[i];

			// 属性追加
			pXMLNode->appendChild(pXMLElement2);

			// 要素の文字列取得
			_bstr_t xml = pXMLNode->Getxml();
			printf("%s\n", (char*)xml);

			pXMLNode.Release();
			pXMLNode = NULL;
		}

		pXMLNodes.Release();
		pXMLNodes = NULL;

		pXMLElement2.Release();
		pXMLElement2 = NULL;
	}

	pXMLElement.Release();
	pXMLElement = NULL;
	pXMLDoc.Release();
	pXMLDoc = NULL;

	::CoUninitialize();

	return 0;
}

