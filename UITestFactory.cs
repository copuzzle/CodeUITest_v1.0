using System;
using System.Drawing;
using System.Windows.Input;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
using System.Data;
using System.Data.Odbc;
using System.Collections.Generic;

namespace 单元测试_Entity.MVCEntity
{
    /// <summary>
    /// IE document对象的操作
    /// </summary>
    internal class MyUIFactory
    {
        /// <summary>
        ///  静态浏览器，只运行一个浏览器
        /// </summary>
        static BrowserWindow browserInstance;
        static BrowserControler browserControler;
        public static BrowserControler BrowserView
        {
            get
            {
                if (browserControler == null)
                {
                    browserControler = new BrowserControler();
                }
                return browserControler;
            }
        }
        public static BrowserWindow Browser
        {
            get
            {
                if (browserInstance == null)
                {
                    browserInstance = new BrowserWindow();
                    browserInstance.CopyFrom(BrowserWindow.Launch());
                }
                return browserInstance;
            }
        }
        public static void NavigateToUrl(string url)
        {
            Browser.NavigateToUrl(new Uri(url));
        }
        public static void NavigateToRelativeUrl(string url)
        {
            BrowserWindow browser  = Browser;
            Browser.NavigateToUrl(new Uri(new Uri(url,UriKind.Relative), browser.Uri));
        }
        public static void Close()
        {
            browserInstance.Close();
            browserInstance = null;
            browserControler = null;
        }
    }

    class BrowserControler
    {
        BrowserWindow Browser { get { return MyUIFactory.Browser; } }
        public HtmlDocument Doc { get { return new HtmlDocument(Browser); } }
        public IntPtr WindowHandle { get { return Browser.WindowHandle; } }
        public object ActiveX
        {
            get
            {
                //return ((mshtml.HTMLDocumentClass)((mshtml.HTMLBodyClass)(((Microsoft.VisualStudio.TestTools.UITesting.UITestControl)(Doc)).NativeElement)).document).parentWindow; 
                object win = ((object[])(((new WinWindow(MyUIFactory.Browser)).NativeElement)))[0];
                return win;
                //return (mshtml.HTMLBodyClass)(((Microsoft.VisualStudio.TestTools.UITesting.UITestControl)(Doc)).NativeElement); 
            }
        }

        public void CaptureImage(string path)
        {
            Image img = Doc.CaptureImage();
            //string path = "c:\\qq"+s+".bmp";
            img.Save(path);
        }
        public string GetUrl()
        {
            return Browser.Uri.ToString();
        }

        public void TextBox(string prefix, string id, string value)
        {
            HtmlEdit TextBox = new HtmlEdit(Doc);
            TextBox.SearchProperties[HtmlEdit.PropertyNames.Id] = prefix + id;
            //string html = (((mshtml.HTMLBodyClass)(((Microsoft.VisualStudio.TestTools.UITesting.UITestControl)(Doc))
            //.NativeElement)).parentElement).innerHTML;
            TextBox.Text = value;
        }

        public void TextArea(string prefix, string id, string value)
        {
            HtmlTextArea TextBox = new HtmlTextArea(Doc);
            TextBox.SearchProperties[HtmlTextArea.PropertyNames.Id] = prefix + id;
            TextBox.Text = value;
        }


        public void RadioButton(string prefix, string id, bool value)
        {
            HtmlRadioButton RadioButton = new HtmlRadioButton(Doc);
            RadioButton.SearchProperties[HtmlRadioButton.PropertyNames.Id] = prefix + id;
            RadioButton.Selected = value;
        }
        public void RadioButton(string value)
        {
            RadioButton(value, true);
        }
        public void RadioButton(string value, bool selected)
        {
            HtmlRadioButton rb = new HtmlRadioButton(Doc);
            rb.SearchProperties[HtmlRadioButton.PropertyNames.Value] = value;
            rb.Selected = selected;
        }
        public bool CellCheck(string expect)
        {
            HtmlCell hc = new HtmlCell(Doc);
            //except的内容必需是cell的全文字段
            hc.SearchProperties[HtmlCell.PropertyNames.InnerText] = expect;
            //hc.SearchProperties[HtmlCell.PropertyNames.InnerText].Contains(expect);

            if (hc.InnerText.Contains(expect))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public void CheckBox(string value , bool selected) 
        {
            HtmlCheckBox cb = new HtmlCheckBox(Doc) ; 
            cb.SearchProperties[HtmlCheckBox.PropertyNames.Value] = value ;
            cb.Checked = selected;
        }
        public void CheckBox(string prefix, string id, bool value)
        {
            HtmlCheckBox CheckBox = new HtmlCheckBox(Doc);
            CheckBox.SearchProperties[HtmlRadioButton.PropertyNames.Id] = prefix + id;
            CheckBox.Checked = value;
        }
        public void DropDownList(string prefix, string id, string value)
        {
            HtmlComboBox DropDownList = new HtmlComboBox(Doc);
            DropDownList.SearchProperties[HtmlRadioButton.PropertyNames.Id] = prefix + id;
            DropDownList.SelectedItem = value;
        }
        public void ButtonClick(string prefix, string id)
        {
            HtmlInputButton Button = new HtmlInputButton(Doc);
            Button.SearchProperties[HtmlInputButton.PropertyNames.Id] = prefix + id;
            Mouse.Click(Button);
        }
        public void ButtonClick(string prefix, string id, int num = 1)
        {
            HtmlInputButton btn = new HtmlInputButton(Doc);
            btn.SearchProperties[HtmlInputButton.PropertyNames.Id] = prefix + id;
            for (int i = 0; i < num; i++)
            {
                Mouse.Click(btn);
            }
        }
        public bool ButtonExist(string prefix, string id)
        {
            HtmlInputButton Button = new HtmlInputButton(Doc);
            Button.SearchProperties[HtmlInputButton.PropertyNames.Id] = prefix + id;
            if (Button.Id != "")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void DivClick(string id)
        {
            HtmlDiv Div = new HtmlDiv(Doc);
            Div.SearchProperties[HtmlDiv.PropertyNames.Id] = id;
            Mouse.Click(Div);
        }
        public void DivClick(string id, int num = 1)
        {
            HtmlDiv Div = new HtmlDiv(Doc);
            Div.SearchProperties[HtmlDiv.PropertyNames.Id] = id;
            for (int i = 0; i < num; i++)
            {
                Mouse.Click(Div);
            }
        }
        public void ImageClick(string id)
        {
            HtmlImage Image = new HtmlImage(Doc);
            Image.SearchProperties[HtmlImage.PropertyNames.Id] = id;
            Mouse.Click(Image);
        }
        //public void HyperLinkClick(string tagid, string tagName, string link)
        public void HyperLinkClick(string parenttagid, string link)
        {
            ////UINVCustom:HtmlCustom ///////////
            HtmlCustom UIN = new HtmlCustom(Doc);
            UIN.SearchProperties["Id"] = parenttagid;
            //UIN.SearchProperties["TagName"] = tagName;
            HtmlHyperlink HLink = new HtmlHyperlink(UIN);
            HLink.FilterProperties[HtmlHyperlink.PropertyNames.Id] = link;
            Mouse.Click(HLink);
        }
        //改进相对url问题
        public void HyperLinkClick(string href)
        {
            HtmlHyperlink Hyperlink = new HtmlHyperlink(Doc);
            string urlhead = href.Substring(4);
            if (urlhead=="http")
            {
                Hyperlink.SearchProperties[HtmlHyperlink.PropertyNames.Href] = href;
            }
            else
            {
                Hyperlink.SearchProperties[HtmlHyperlink.PropertyNames.AbsolutePath] = href;
            }
            Mouse.Click(Hyperlink);
        }
        public void HyperLinkIdClick(string prefix, string id)
        {
            string newid = prefix + id;
            HyperLinkIdClick(newid);
        }
        public void HyperLinkIdClick(string id)
        {
            HtmlHyperlink Hyperlink = new HtmlHyperlink(Doc);
            Hyperlink.SearchProperties[HtmlHyperlink.PropertyNames.Id] = id;
            Mouse.Click(Hyperlink);
        }
        public void HyperlinkTextClick(string innertext)
        {
            HtmlHyperlink a = new HtmlHyperlink(Doc);
            a.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = innertext;
            Mouse.Click(a);

        }
        public void TableClick(string tagid, string Buttonid)
        {
            HtmlCustom UIN = new HtmlCustom(Doc);
            UIN.SearchProperties["Id"] = tagid;
            HtmlTable button = new HtmlTable(UIN);
            button.FilterProperties[HtmlTable.PropertyNames.Id] = Buttonid;
            Mouse.Click(button);
        }
        public void TableClick(string id)
        {
            HtmlTable Table = new HtmlTable(Doc);
            Table.SearchProperties[HtmlTable.PropertyNames.Id] = id;
            Mouse.Click(Table);
        }
        public void TdClick(string id)
        {
            HtmlCell Td = new HtmlCell(Doc);
            Td.SearchProperties[HtmlCell.PropertyNames.Id] = id;
            Mouse.Click(Td);
        }

        public string GetTextBoxValue(string id)
        {
            HtmlEdit TextBox = new HtmlEdit(Doc);
            TextBox.SearchProperties[HtmlEdit.PropertyNames.Id] = id;
            return TextBox.Text;
        }
        public string GetLabelValue(string id)
        {
            HtmlLabel Label = new HtmlLabel(Doc);
            Label.SearchProperties[HtmlLabel.PropertyNames.Id] = id;
            return Label.InnerText;
        }

        public void SendKeys(string keys)
        {
            HtmlDocument doc = new HtmlDocument(Doc);
            Keyboard.SendKeys(keys);
        }
        public void SendWinAndKeys(string keys)
        {
            HtmlDocument doc = new HtmlDocument(Doc);
            Keyboard.SendKeys(keys, ModifierKeys.Windows);
        }
    }
    class BrowserWinContoler
    {
        BrowserWindow browser;
        public BrowserWinContoler(string name)
        {
            browser = new BrowserWindow();
            browser.SearchProperties[BrowserWindow.PropertyNames.Name] = name;
        }
        /// <summary>
        /// 获取窗口某个控件，并点击。适用于控件在窗口中唯一存在
        /// </summary>
        /// <param name="controltype">控件类型</param>
        /// 如梦:"DropDownButton"
        public void WindowControlTypeClick(string controltype)
        {
            //WinSplitButton
            //UIItemWindow

            WinSplitButton wsb = new WinSplitButton(browser);
            WinControl wc = new WinControl(wsb);
            wc.SearchProperties[WinControl.PropertyNames.ControlType] = controltype;
            Mouse.Click(wc);
        }
       

    }

}
