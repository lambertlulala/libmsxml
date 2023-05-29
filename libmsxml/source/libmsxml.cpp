#include "libmsxml.h"

#include <objbase.h>  
#include <msxml6.h>

#include <stdexcept>

WinMSXML::WinMSXML() : m_init{ false }
{

}

WinMSXML::WinMSXML(WinMSXML&& rhs) noexcept
{
    m_xmlDoc = rhs.m_xmlDoc.Detach();
}

bool WinMSXML::Init()
{
    if (!m_init && SUCCEEDED(m_xmlDoc.CoCreateInstance(__uuidof(DOMDocument60), nullptr, CLSCTX_INPROC_SERVER)))
    {
        m_xmlDoc->put_async(VARIANT_FALSE);
        m_xmlDoc->put_validateOnParse(VARIANT_FALSE);
        m_xmlDoc->put_resolveExternals(VARIANT_FALSE);
        m_init = true;
    }

    return m_init;
}

void WinMSXML::Load(const std::wstring& xmlFile)
{
    if (!IsInit())
        throw std::runtime_error("Unable to initialize.");

    CComVariant xmlSrc = xmlFile.c_str();
    VARIANT_BOOL varStatus;
    m_xmlDoc->load(xmlSrc, &varStatus);
    if (VARIANT_TRUE != varStatus)
        throw std::runtime_error("Unable to load file.");
}

void WinMSXML::LoadFromString(const std::wstring& xmlData)
{
    if (!IsInit())
        throw std::runtime_error("Unable to initialize.");

    VARIANT_BOOL varStatus;
    m_xmlDoc->loadXML(CComBSTR(xmlData.c_str()), &varStatus);
    if (VARIANT_TRUE != varStatus)
        throw std::runtime_error("Unable to load from string.");
}

WinMSXML::XMLElementList WinMSXML::GetElementsByTagName(const std::wstring& tagName)
{
    if (!IsInit())
        throw std::runtime_error("Not initialized.");

    XMLElementList elementList;
    if (FAILED(m_xmlDoc->getElementsByTagName(CComBSTR(tagName.c_str()), &elementList))
        || !elementList)
        throw std::runtime_error("Cannot get elements by tag name.");

    return elementList;
}

WinMSXML::XMLElementList WinMSXML::GetChildElements(const XMLElement &element)
{
    if (!element)
        throw std::invalid_argument("Invalid element.");

    XMLElementList childElementList;
    if (FAILED(element->get_childNodes(&childElementList))
        || !childElementList)
        throw std::runtime_error("Cannot get child nodes.");

    return childElementList;
}

WinMSXML::XMLAttributes WinMSXML::GetElementAttributes(const XMLElement& element)
{
    if (!element)
        throw std::invalid_argument("Invalid element.");

    XMLAttributes attributes;
    if (FAILED(element->get_attributes(&attributes))
        || !attributes)
        throw std::runtime_error("Cannot get attributes.");

    return attributes;
}

const std::wstring WinMSXML::GetAttributeValue(const XMLElement& element, const std::wstring& attributeName)
{
    std::wstring value;

    try
    {
        auto attributes = GetElementAttributes(element);
        WinMSXML::XMLElement namedNode;
        if (FAILED(attributes->getNamedItem(CComBSTR(attributeName.c_str()), &namedNode))
            || !namedNode)
            throw std::runtime_error("Cannot get named item.");

        CComBSTR text;
        namedNode->get_text(&text);
        value = text;
    }
    catch (std::exception& e)
    {
        throw;
    }

    return value;
}
