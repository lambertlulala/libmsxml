#include "pch.h"

#include "libmsxml.h"

#include <Windows.h>
#include <string>
#include <format>
#include <fstream>
#include <vector>

TEST(TestLoadXML, TestLoadFromString)
{
    ASSERT_TRUE(SUCCEEDED(CoInitializeEx(nullptr, COINIT_MULTITHREADED)));

    {
        WCHAR dir[256]{};
        GetCurrentDirectory(_countof(dir), dir);
        std::wfstream f(std::format(L"{}\\..\\testdata\\stocks.xml", dir), std::iostream::in);
        ASSERT_TRUE(f);

        f.seekg(f.beg, f.end);
        size_t size = f.tellg();
        f.seekg(0);
        std::vector<wchar_t> buffer(size);
        f.read(buffer.data(), size);
        try
        {
            WinMSXML xml;
            ASSERT_TRUE(xml.Init());
            auto xmlDoc = xml.GetXMLDocument();
            std::wstring xmlText(buffer.data(), buffer.size());
            VARIANT_BOOL success;
            ASSERT_TRUE(SUCCEEDED(xmlDoc->loadXML(CComBSTR(xmlText.c_str()), &success)));
            ASSERT_TRUE(SUCCEEDED(xmlDoc->save(CComVariant(L"test.xml"))));
        }
        catch (const std::exception& e)
        {
            std::cout << e.what() << std::endl;
            ASSERT_FALSE(true);
        }
    }

    CoUninitialize();
}

TEST(TestLoadXML, TestLoadFile)
{
    ASSERT_TRUE(SUCCEEDED(CoInitializeEx(nullptr, COINIT_MULTITHREADED)));

    {
        WCHAR dir[256]{};
        GetCurrentDirectory(_countof(dir), dir);

        try
        {
            WinMSXML xml;
            EXPECT_TRUE(xml.Init());
            xml.Load(std::format(L"{}\\..\\testdata\\nonexisting_stocks.xml", dir));
            ASSERT_FALSE(true);
        }
        catch (const std::exception& e)
        {
            ASSERT_TRUE(true);
        }

        try
        {
            WinMSXML xml;
            EXPECT_TRUE(xml.Init());
            xml.Load(std::format(L"{}\\..\\testdata\\stocks.xml", dir));
            ASSERT_TRUE(true);

            auto xmlDoc = xml.GetXMLDocument();
            CComBSTR xmlText;
            xmlDoc->get_xml(&xmlText);
            ASSERT_TRUE(xmlText.Length() > 0);
        }
        catch (const std::exception& e)
        {
            std::cout << e.what() << std::endl;
            ASSERT_FALSE(true);
        }
    }

    CoUninitialize();
}

TEST(TestGetElements, TestGetElemtsByTag)
{
    ASSERT_TRUE(SUCCEEDED(CoInitializeEx(nullptr, COINIT_MULTITHREADED)));

    {
        WCHAR dir[256]{};
        GetCurrentDirectory(_countof(dir), dir);

        try
        {
            WinMSXML xml;
            EXPECT_TRUE(xml.Init());

            xml.Load(std::format(L"{}\\..\\testdata\\stocks.xml", dir));
            ASSERT_TRUE(true);

            auto elemts = xml.GetElementsByTagName(L"stock");
            long length;
            elemts->get_length(&length);
            EXPECT_TRUE(length > 0);
        }
        catch (const std::exception& e)
        {
            std::cout << e.what() << std::endl;
            ASSERT_FALSE(true);
        }
    }

    CoUninitialize();
}

TEST(TestGetElements, TestGetChildElemts)
{
    ASSERT_TRUE(SUCCEEDED(CoInitializeEx(nullptr, COINIT_MULTITHREADED)));

    {
        WCHAR dir[256]{};
        GetCurrentDirectory(_countof(dir), dir);

        try
        {
            WinMSXML xml;
            EXPECT_TRUE(xml.Init());

            xml.Load(std::format(L"{}\\..\\testdata\\stocks.xml", dir));
            ASSERT_TRUE(true);

            auto elements = xml.GetElementsByTagName(L"stock");
            long length;
            elements->get_length(&length);
            EXPECT_TRUE(length > 0);
            elements->reset();
            for (long i = 0; i < length; i++)
            {
                WinMSXML::XMLElement element;
                ASSERT_TRUE(SUCCEEDED(elements->get_item(i, &element)));

                auto childElements = xml.GetChildElements(element);
                WinMSXML::XMLElement firstChildElement;
                ASSERT_TRUE(SUCCEEDED(childElements->get_item(0, &firstChildElement)));

                CComBSTR text;
                ASSERT_TRUE(SUCCEEDED(firstChildElement->get_text(&text)));
                std::wstring firstChildText(text);
                switch (i)
                {
                case 0:
                    ASSERT_TRUE(L"new" == firstChildText);
                    break;

                case 1:
                    ASSERT_TRUE(L"zacx corp" == firstChildText);
                    break;

                case 2:
                    ASSERT_TRUE(L"zaffymat inc" == firstChildText);
                    break;

                case 3:
                    ASSERT_TRUE(L"zysmergy inc" == firstChildText);
                    break;

                default:
                    ASSERT_FALSE(true);
                    break;
                }
            }
        }
        catch (const std::exception& e)
        {
            std::cout << e.what() << std::endl;
            ASSERT_FALSE(true);
        }
    }

    CoUninitialize();
}

TEST(TestGetElements, TestGetElementAttribute)
{
    ASSERT_TRUE(SUCCEEDED(CoInitializeEx(nullptr, COINIT_MULTITHREADED)));

    {
        WCHAR dir[256]{};
        GetCurrentDirectory(_countof(dir), dir);

        try
        {
            WinMSXML xml;
            EXPECT_TRUE(xml.Init());

            xml.Load(std::format(L"{}\\..\\testdata\\stocks.xml", dir));
            ASSERT_TRUE(true);

            WinMSXML::XMLElement firstChildElement;
            auto elements = xml.GetElementsByTagName(L"stock");
            ASSERT_TRUE(SUCCEEDED(elements->get_item(0, &firstChildElement)));

            auto value = xml.GetAttributeValue(firstChildElement, L"exchange");
            ASSERT_TRUE(L"nasdaq" == value);
        }
        catch (const std::exception& e)
        {
            std::cout << e.what() << std::endl;
            ASSERT_FALSE(true);
        }
    }

    CoUninitialize();
}

TEST(TestGetElements, TestGetFromBuffer)
{
    ASSERT_TRUE(SUCCEEDED(CoInitializeEx(nullptr, COINIT_MULTITHREADED)));

    {
        WCHAR dir[256]{};
        GetCurrentDirectory(_countof(dir), dir);
        std::wfstream f(std::format(L"{}\\..\\testdata\\stocks.xml", dir), std::iostream::in);
        if (f)
        {
            f.seekg(f.beg, f.end);
            size_t size = f.tellg();
            f.seekg(0);

            std::wstring xmlString(size, 0);
            f.read(const_cast<wchar_t *>(xmlString.data()), size);

            WinMSXML xml;
            EXPECT_TRUE(xml.Init());
            xml.LoadFromString(xmlString);
            auto xmlDoc = xml.GetXMLDocument();
            xmlDoc->save(CComVariant(L"test.xml"));
        }
    }

    CoUninitialize();
}
