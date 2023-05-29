#pragma once

#include <atlbase.h>
#include <string>

class WinMSXML
{
private:
    CComPtr<IXMLDOMDocument> m_xmlDoc;
    bool m_init;

public:
    using XMLElement = CComPtr<IXMLDOMNode>;
    using XMLElementList = CComPtr<IXMLDOMNodeList>;
    using XMLAttributes = CComPtr<IXMLDOMNamedNodeMap>;
    using XMLDocument = CComPtr<IXMLDOMDocument>;

    WinMSXML();
    WinMSXML(const WinMSXML& rhs) = delete;
    WinMSXML(WinMSXML&& rhs) noexcept;
    void operator= (WinMSXML& rhs) = delete;

    virtual bool Init();
    const bool IsInit() const { return m_init; }
    virtual void Load(const std::wstring& xmlFile);
    virtual void LoadFromString(const std::wstring& xmlData);

    XMLDocument GetXMLDocument() const { return m_xmlDoc; }
    virtual XMLElementList GetElementsByTagName(const std::wstring& tagName);
    virtual XMLElementList GetChildElements(const XMLElement &element);
    virtual XMLAttributes GetElementAttributes(const XMLElement &element);
    virtual const std::wstring GetAttributeValue(const XMLElement &element, const std::wstring& attributeName);
};
