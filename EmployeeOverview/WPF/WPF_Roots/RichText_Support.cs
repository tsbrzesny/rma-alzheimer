using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;

namespace WPF_Roots
{
    public class RTB_Support
    {

        private FlowDocument flowDoc = null;
        //
        public RTB_Support(RichTextBox target_rtb = null)
        {
            if (target_rtb != null)
                flowDoc = target_rtb.Document;
            else
                flowDoc = new FlowDocument();
        }

        public FlowDocument Document
        { get { return flowDoc; } }

        public BlockCollection Blocks
        { get { return flowDoc.Blocks; } }

        public void Clear()
        {
            Blocks.Clear();
        }

        public void AddBlock(Block newBlock)
        {
            Blocks.Add(newBlock);
        }

        private ThicknessConverter tConv = new ThicknessConverter();
        public void AddLine(string text, String margin = "0")
        {
            Paragraph pr = new Paragraph() { Margin = (Thickness)tConv.ConvertFromString(margin) };
            pr.Inlines.Add(text);
            Blocks.Add(pr);
        }

        public void AddPara(string content, String margin = "0")
        {
            var pr = CreatePara(content, margin);
            Blocks.Add(pr);
        }

        public void AddTable(string content, String margin = "0")
        {
            var tbl = CreateTable(content, margin);
            Blocks.Add(tbl);
        }


        public static TextElement CreateDocElement(string elementType, string content, String margin, string rootAttr)
        {
            var pc = string.Format(@"<{0}
                xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
                xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml' Margin=""{1}"" {2}> {3} </{0}>", elementType, margin, rootAttr, content);
            var te = (TextElement)XamlReader.Parse(pc);
            return te;
        }

        public static Inline CreateInline(string content)
        {
            var m= Regex.Match(content, @"<\w+");
            if (m.Success)
            {
                content = content.Insert(m.Index + m.Length, @" xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
                                                                xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'");
            }
            var inline = (Inline)XamlReader.Parse(content);
            return inline;
        }

        public static List<Inline> CreateInlines(string xaml)
        {
            var dummyPara = CreatePara(xaml);
            return new List<Inline>(dummyPara.Inlines);
        }


        public Table CreateTable(string content, String margin = "0", string rootAttr = "")
        {
            return (Table)CreateDocElement("Table", content, margin, rootAttr);
        }

        public static Paragraph CreatePara(string content, String margin = "0", string rootAttr = "")
        {
            return (Paragraph)CreateDocElement("Paragraph", content, margin, rootAttr);
        }

        public List CreateList(string content, String margin = "0", string rootAttr = "")
        {
            return (List)CreateDocElement("List", content, margin, rootAttr);
        }

        public Section CreateSection(string content, String margin = "0", string rootAttr = "")
        {
            return (Section)CreateDocElement("Section", content, margin, rootAttr);
        }

        public void ReplaceOrAdd(Block block, Block replaceThisBlock = null)
        {
            if (block == null)
                return;

            if (replaceThisBlock == null)
            {
                // replace the first block with the same name
                var firstMatchingBlock = Blocks.FirstOrDefault(_b => _b.Name == block.Name);
                if (firstMatchingBlock != null)
                {
                    Blocks.InsertAfter(firstMatchingBlock, block);
                    Blocks.Remove(firstMatchingBlock);
                    return;
                }
            }

            else if (Blocks.Contains(replaceThisBlock))
            {
                Blocks.InsertAfter(replaceThisBlock, block);
                Blocks.Remove(replaceThisBlock);
                return;
            }

            // no replace action is poassible - add as new block
            Blocks.Add(block);
        }

        public static string XAMLEsc(string s)
        {
            s = s.Replace("&", "&amp;");
            s = s.Replace("<", "&lt;");
            s = s.Replace(">", "&gt;");
            s = s.Replace(@"""", "&quot;");
            s = Regex.Replace(s, @"\n\r?|\\r\n?", "<LineBreak />");
            return s;
        }

    }


    public static class FlowDocument_Extensions
    {
        public static void ReplaceByName(this Paragraph parent, string xaml)
        {
            var newInline = RTB_Support.CreateInline(xaml);
            var oldInline = parent.Inlines.FirstOrDefault(_i => _i.Name == newInline.Name);
            if (oldInline != null)
            {
                parent.Inlines.InsertAfter(oldInline, newInline);
                parent.Inlines.Remove(oldInline);
            }
            else
                parent.Inlines.Add(newInline);
        }

        public static void AddPara(this Section parent, string content, String margin = "0", string rootAttr = "")
        {
            var pr = RTB_Support.CreatePara(content, margin, rootAttr);
            parent.Blocks.Add(pr);
        }

    }
}
