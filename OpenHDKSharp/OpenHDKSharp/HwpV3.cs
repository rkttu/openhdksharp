/*
 * HwpV3.cs
 *
 * Copyright (C) 2013 Junghyun Nam <rkttu@rkttu.com>
 * Based on Hodong Kim's libhwp project. <cogniti@gmail.com>
 * https://github.com/cogniti
 * 
 * This library is free software: you can redistribute it and/or modify it
 * under the terms of the GNU Lesser General Public License as published
 * by the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This library is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.IO.Compression;

namespace OpenHDKSharp.V3
{
    public static class HwpV3BinaryReaderExtension
    {
        public static ushort ReadWord(this BinaryReader reader)
        { return reader.ReadUInt16(); }

        public static short ReadSWord(this BinaryReader reader)
        { return reader.ReadInt16(); }

        public static uint ReadDWord(this BinaryReader reader)
        { return reader.ReadUInt32(); }

        public static int ReadSDWord(this BinaryReader reader)
        { return reader.ReadInt32(); }

        public static ushort ReadHChar(this BinaryReader reader)
        { return reader.ReadUInt16(); }

        public static sbyte ReadEChar(this BinaryReader reader)
        { return reader.ReadSByte(); }

        public static byte ReadKChar(this BinaryReader reader)
        { return reader.ReadByte(); }

        public static ushort ReadHUnit(this BinaryReader reader)
        { return reader.ReadUInt16(); }

        public static short ReadSHUnit(this BinaryReader reader)
        { return reader.ReadInt16(); }

        public static uint ReadHUnit32(this BinaryReader reader)
        { return reader.ReadUInt32(); }

        public static int ReadSHUnit32(this BinaryReader reader)
        { return reader.ReadInt32(); }

        public static byte[] ReadKCharArray(this BinaryReader reader, int count)
        { return reader.ReadBytes(count); }

        public static ushort[] ReadWordArray(this BinaryReader reader, int count)
        {
            ushort[] results = new ushort[count];
            for (int i = 0; i < count; i++)
                results[i] = reader.ReadWord();
            return results;
        }

        public static ushort[] ReadHUnitArray(this BinaryReader reader, int count)
        {
            ushort[] results = new ushort[count];
            for (int i = 0; i < count; i++)
                results[i] = reader.ReadWord();
            return results;
        }

        public static ushort[] ReadHCharArray(this BinaryReader reader, int count)
        {
            ushort[] results = new ushort[count];
            for (int i = 0; i < count; i++)
                results[i] = reader.ReadWord();
            return results;
        }

        public static string ReadKCharString(this BinaryReader reader, int count)
        {
            return Encoding.GetEncoding(1361).GetString(ReadKCharArray(reader, count).TakeWhile(x => x != 0x00).ToArray());
        }

        public static string ReadHCharString(this BinaryReader reader, int count)
        {
            return Hnc2Unicode.hnstring_to_unicodestr(ReadHCharArray(reader, count));
        }

        public static sbyte[] ReadSBytes(this BinaryReader reader, int count)
        {
            sbyte[] results = new sbyte[count];
            for (int i = 0; i < count; i++)
                results[i] = reader.ReadSByte();
            return results;
        }
    }

    public sealed class HwpV3Document
    {
        public static HwpV3Document Load(string filePath)
        {
            HwpV3Document doc = new HwpV3Document();

            using (Stream file = File.OpenRead(filePath))
            using (BinaryReader reader = new BinaryReader(file))
            {
                doc.Signature = HwpV3DocumentSignature.Create(reader);
                doc.Information = HwpV3DocumentInformation.Create(reader);
                doc.Summary = HwpV3DocumentSummary.Create(reader);
                doc.InformationBlocks = HwpV3DocumentInformationBlockCollection.Create(reader, doc.Information);

                bool requireDecompress = ((int)doc.Information.Compressed != 0);
                MemoryStream extractedFile = new MemoryStream();

                byte[] buffer = new byte[64000];
                int read = 0;

                if (requireDecompress)
                {
                    using (DeflateStream deflateStream = new DeflateStream(file, CompressionMode.Decompress, true))
                    {
                        while ((read = deflateStream.Read(buffer, 0, buffer.Length)) > 0)
                            extractedFile.Write(buffer, 0, read);
                    }
                }
                else
                {
                    while ((read = file.Read(buffer, 0, buffer.Length)) > 0)
                        extractedFile.Write(buffer, 0, read);
                }

                extractedFile.Seek(0L, SeekOrigin.Begin);

                using (extractedFile)
                using (BinaryReader subReader = new BinaryReader(extractedFile))
                {
                    doc.FontNames = HwpV3FontNames.Create(subReader);
                    doc.Styles = HwpV3DocumentStyleCollection.Create(subReader);
                    doc.Paragraph = HwpV3Paragraph.Create(subReader);
                    doc.AdditionalBlock1 = HwpV3AdditionalBlock.Create(subReader);

                    uint additionalBlock2Size = 0u;
                    bool hasAdditionalBlock2 = false;
                    BinaryReader target = (requireDecompress ? reader : subReader);
                    target.BaseStream.Seek(-8, SeekOrigin.End);
                    uint id = target.ReadDWord();
                    if (id == 0u)
                    {
                        additionalBlock2Size = target.ReadDWord();
                        if (additionalBlock2Size != 0u)
                            hasAdditionalBlock2 = true;
                    }

                    if (hasAdditionalBlock2)
                    {
                        target.BaseStream.Position = target.BaseStream.Length - additionalBlock2Size;
                        doc.AdditionalBlock2 = HwpV3AdditionalBlock.Create(reader);
                    }
                }

                return doc;
            }
        }

        public HwpV3DocumentSignature Signature { get; set; }

        public HwpV3DocumentInformation Information { get; set; }

        public HwpV3DocumentSummary Summary { get; set; }

        public HwpV3DocumentInformationBlockCollection InformationBlocks { get; set; }

        public HwpV3FontNames FontNames { get; set; }

        public HwpV3DocumentStyleCollection Styles { get; set; }

        public HwpV3Paragraph Paragraph { get; set; }

        public HwpV3AdditionalBlock AdditionalBlock1 { get; set; }

        public HwpV3AdditionalBlock AdditionalBlock2 { get; set; }
    }

    public sealed class HwpV3DocumentSignature
    {
        // 파일 시그니처
        public byte[] Signature;

        public static HwpV3DocumentSignature Create(BinaryReader reader)
        {
            byte[] signature = reader.ReadBytes(30);
            bool result = (
                signature != null && signature.Length == 30 &&
                signature[0x00] == 0x48 && signature[0x01] == 0x57 && signature[0x02] == 0x50 &&
                signature[0x03] == 0x20 && signature[0x04] == 0x44 && signature[0x05] == 0x6F &&
                signature[0x06] == 0x63 && signature[0x07] == 0x75 && signature[0x08] == 0x6D &&
                signature[0x09] == 0x65 && signature[0x0A] == 0x6E && signature[0x0B] == 0x74 &&
                signature[0x0C] == 0x20 && signature[0x0D] == 0x46 && signature[0x0E] == 0x69 &&
                signature[0x0F] == 0x6C && signature[0x10] == 0x65 && signature[0x11] == 0x20 &&
                signature[0x12] == 0x56 && signature[0x13] == 0x33 && signature[0x14] == 0x2E &&
                signature[0x15] == 0x30 && signature[0x16] == 0x30 && signature[0x17] == 0x20 &&
                signature[0x18] == 0x1A && signature[0x19] == 0x01 && signature[0x1A] == 0x02 &&
                signature[0x1B] == 0x03 && signature[0x1C] == 0x04 && signature[0x1D] == 0x05
            );

            if (!result)
                return null;

            HwpV3DocumentSignature item = new HwpV3DocumentSignature();
            item.Signature = signature;
            return item;
        }
    }

    public sealed class HwpV3DocumentInformation
    {
        public static HwpV3DocumentInformation Create(BinaryReader reader)
        {
            HwpV3DocumentInformation result = new HwpV3DocumentInformation();

            // 커서 줄
            result.CursorLine = reader.ReadWord();

            // 커서 칸
            result.CursorColumn = reader.ReadWord();

            // 용지 종류
            result.PaperSize = reader.ReadByte();

            // 용지 방향
            result.PaperOrientation = reader.ReadByte();

            // 용지 길이
            result.PaperHeight = reader.ReadHUnit();

            // 용지 너비
            result.PaperWidth = reader.ReadHUnit();

            // 위쪽 여백
            result.TopMargin = reader.ReadHUnit();

            // 아래쪽 여백
            result.BottomMargin = reader.ReadHUnit();

            // 왼쪽 여백
            result.LeftMargin = reader.ReadHUnit();

            // 오른쪽 여백
            result.RightMargin = reader.ReadHUnit();

            // 머리말 길이
            result.HeaderHeight = reader.ReadHUnit();

            // 꼬리말 길이
            result.FooterHeight = reader.ReadHUnit();

            // 제본 여백
            result.BindingMargin = reader.ReadHUnit();

            // 문서 보호
            result.Protection = reader.ReadDWord();

            // 예약
            result.Reserved = reader.ReadWord();

            // 쪽 번호 연결
            result.ContinuousPageNo = reader.ReadByte();

            // 각주 번호 연결
            result.ContinuousFootnoteNo = reader.ReadByte();

            // 연결 인쇄 파일
            result.LinkedPrintingFile = reader.ReadKCharString(40);

            // 덧붙이는 말
            result.Comments = reader.ReadKCharString(24);

            // 암호 여부
            result.PasswordProtected = reader.ReadWord();

            // 시작 페이지 번호
            result.StartPageNo = reader.ReadWord();

            // 각주 시작 번호
            result.StartFootnoteNo = reader.ReadWord();

            // 각주 개수
            result.FootnoteCount = reader.ReadWord();

            // 각주 분리선과 본문 사이의 간격
            result.FootnoteDividerTopMargin = reader.ReadHUnit();

            // 각주와 본문 사이의 간격
            result.FootnoteTopMargin = reader.ReadHUnit();

            // 각주와 각주 사이의 간격
            result.FootnoteSpacing = reader.ReadHUnit();

            // 각주 번호에 괄호 붙임
            result.FootnoteNumberDecoration = reader.ReadEChar();

            // 각주 분리선 너비
            result.FootnoteDividerWidth = reader.ReadByte();

            // 테두리 간격
            result.BorderSpacing = reader.ReadHUnitArray(4);

            // 테두리 종류
            result.BorderType = reader.ReadWord();

            // 빈 줄 감춤
            result.HideEmptyLine = reader.ReadByte();

            // 틀 옮김
            result.ShapeMoving = reader.ReadByte();

            // 압축
            result.Compressed = reader.ReadByte();

            // Sub Revision
            result.SubRevision = reader.ReadByte();

            // 정보 블록 길이
            result.InformationBlockLength = reader.ReadWord();

            return result;
        }

        public ushort CursorLine { get; set; }

        public ushort CursorColumn { get; set; }

        public byte PaperSize { get; set; }

        public byte PaperOrientation { get; set; }

        public ushort PaperHeight { get; set; }

        public ushort PaperWidth { get; set; }

        public ushort TopMargin { get; set; }

        public ushort BottomMargin { get; set; }

        public ushort LeftMargin { get; set; }

        public ushort RightMargin { get; set; }

        public ushort HeaderHeight { get; set; }

        public ushort FooterHeight { get; set; }

        public ushort BindingMargin { get; set; }

        public uint Protection { get; set; }

        public ushort Reserved { get; set; }

        public byte ContinuousPageNo { get; set; }

        public byte ContinuousFootnoteNo { get; set; }

        public string LinkedPrintingFile { get; set; }

        public string Comments { get; set; }

        public ushort PasswordProtected { get; set; }

        public ushort StartPageNo { get; set; }

        public ushort StartFootnoteNo { get; set; }

        public ushort FootnoteCount { get; set; }

        public ushort FootnoteDividerTopMargin { get; set; }

        public ushort FootnoteTopMargin { get; set; }

        public ushort FootnoteSpacing { get; set; }

        public sbyte FootnoteNumberDecoration { get; set; }

        public byte FootnoteDividerWidth { get; set; }

        public ushort[] BorderSpacing { get; set; }

        public ushort BorderType { get; set; }

        public byte HideEmptyLine { get; set; }

        public byte ShapeMoving { get; set; }

        public byte Compressed { get; set; }

        public byte SubRevision { get; set; }

        public ushort InformationBlockLength { get; set; }
    }

    public sealed class HwpV3DocumentSummary
    {
        public static HwpV3DocumentSummary Create(BinaryReader reader)
        {
            HwpV3DocumentSummary result = new HwpV3DocumentSummary();

            // 제목
            result.Title = reader.ReadHCharString(56);

            // 주제
            result.Subject = reader.ReadHCharString(56);

            // 지은이
            result.Author = reader.ReadHCharString(56);

            // 날짜
            result.Date = reader.ReadHCharString(56);

            // 키워드 1, 2
            result.Keywords = new List<string>();
            for (int i = 0; i < 2; i++)
            {
                string temp = String.Concat(reader.ReadHCharString(56)).Trim();

                if (temp.Length < 1)
                    continue;

                result.Keywords.Add(temp);
            }

            // 기타 1, 2, 3
            result.Misc = new List<string>();
            for (int i = 0; i < 3; i++)
            {
                string temp = String.Concat(reader.ReadHCharString(56)).Trim();

                if (temp.Length < 1)
                    continue;

                result.Misc.Add(temp);
            }

            return result;
        }

        public string Title { get; set; }

        public string Subject { get; set; }

        public string Author { get; set; }

        public string Date { get; set; }

        public List<string> Keywords { get; set; }

        public List<string> Misc { get; set; }

        public DateTime ParsedDate
        { get { return DateTime.Parse(this.Date, CultureInfo.GetCultureInfo("ko-KR")); } }
    }

    public sealed class HwpV3DocumentInformationBlockCollection : List<HwpV3DocumentInformationBlock>
    {
        public static HwpV3DocumentInformationBlockCollection Create(BinaryReader reader, HwpV3DocumentInformation docinfo)
        {
            HwpV3DocumentInformationBlockCollection results = new HwpV3DocumentInformationBlockCollection();
            results.Length = docinfo.InformationBlockLength;

            using (MemoryStream memStream = new MemoryStream(reader.ReadBytes((int)results.Length)))
            using (BinaryReader subReader = new BinaryReader(memStream))
            {
                while (memStream.Position < memStream.Length)
                    results.Add(HwpV3DocumentInformationBlock.Create(subReader));
            }

            return results;
        }

        public ushort Length { get; set; }
    }

    public sealed class HwpV3DocumentInformationBlock
    {
        public static HwpV3DocumentInformationBlock Create(BinaryReader reader)
        {
            HwpV3DocumentInformationBlock block = new HwpV3DocumentInformationBlock();

            block.ID = reader.ReadWord();
            block.Length = reader.ReadWord();
            block.Content = reader.ReadBytes((int)block.Length);

            return block;
        }

        public ushort ID { get; set; }

        public ushort Length { get; set; }

        public byte[] Content { get; set; }
    }

    public sealed class HwpV3FontNames
    {
        public static HwpV3FontNames Create(BinaryReader reader)
        {
            HwpV3FontNames result = new HwpV3FontNames();
            List<string> items = null;
            ushort fontCount = 0;

            fontCount = reader.ReadWord();
            items = new List<string>((int)fontCount);
            for (ushort i = 0; i < fontCount; i++)
                items.Add(reader.ReadKCharString(40));
            result.KoreanFonts = items;

            fontCount = reader.ReadWord();
            items = new List<string>((int)fontCount);
            for (ushort i = 0; i < fontCount; i++)
                items.Add(reader.ReadKCharString(40));
            result.EnglishFonts = items;

            fontCount = reader.ReadWord();
            items = new List<string>((int)fontCount);
            for (ushort i = 0; i < fontCount; i++)
                items.Add(reader.ReadKCharString(40));
            result.ChineseCharacterFonts = items;

            fontCount = reader.ReadWord();
            items = new List<string>((int)fontCount);
            for (ushort i = 0; i < fontCount; i++)
                items.Add(reader.ReadKCharString(40));
            result.JapaneseFonts = items;

            fontCount = reader.ReadWord();
            items = new List<string>((int)fontCount);
            for (ushort i = 0; i < fontCount; i++)
                items.Add(reader.ReadKCharString(40));
            result.OtherFonts = items;

            fontCount = reader.ReadWord();
            items = new List<string>((int)fontCount);
            for (ushort i = 0; i < fontCount; i++)
                items.Add(reader.ReadKCharString(40));
            result.SymbolFonts = items;

            fontCount = reader.ReadWord();
            items = new List<string>((int)fontCount);
            for (ushort i = 0; i < fontCount; i++)
                items.Add(reader.ReadKCharString(40));
            result.CustomFonts = items;

            return result;
        }

        public List<string> KoreanFonts { get; set; }

        public List<string> EnglishFonts { get; set; }

        public List<string> ChineseCharacterFonts { get; set; }

        public List<string> JapaneseFonts { get; set; }

        public List<string> OtherFonts { get; set; }

        public List<string> SymbolFonts { get; set; }

        public List<string> CustomFonts { get; set; }
    }

    public sealed class HwpV3ColumnSettings
    {
        public static HwpV3ColumnSettings Create(BinaryReader reader)
        {
            HwpV3ColumnSettings result = new HwpV3ColumnSettings();

            // 단 수
            result.ColumnCount = reader.ReadByte();

            // 단 구분선
            result.ColumnDivider = reader.ReadByte();

            // 단 간격
            result.ColumnSpacing = reader.ReadHUnit();

            // 예약
            result.Reserved = reader.ReadBytes(4);

            return result;
        }

        public byte ColumnCount { get; set; }

        public byte ColumnDivider { get; set; }

        public ushort ColumnSpacing { get; set; }

        public byte[] Reserved { get; set; }
    }

    public sealed class HwpV3DocumentStyle
    {
        public static HwpV3DocumentStyle Create(BinaryReader reader)
        {
            HwpV3DocumentStyle result = new HwpV3DocumentStyle();

            // 스타일 이름
            result.StyleName = reader.ReadKCharString(20);

            // 글자 모양
            result.FontShape = HwpV3FontShape.Create(reader);

            // 문단 모양
            result.ParagraphShape = HwpV3ParagraphShape.Create(reader);

            return result;
        }

        public string StyleName { get; set; }

        public HwpV3FontShape FontShape { get; set; }

        public HwpV3ParagraphShape ParagraphShape { get; set; }
    }

    public sealed class HwpV3DocumentStyleCollection : List<HwpV3DocumentStyle>
    {
        public static HwpV3DocumentStyleCollection Create(BinaryReader reader)
        {
            HwpV3DocumentStyleCollection results = new HwpV3DocumentStyleCollection();

            ushort styleCount = reader.ReadWord();
            for (ushort i = 0; i < styleCount; i++)
                results.Add(HwpV3DocumentStyle.Create(reader));

            return results;
        }
    }

    public sealed class HwpV3FontShape
    {
        public static HwpV3FontShape Create(BinaryReader reader)
        {
            HwpV3FontShape result = new HwpV3FontShape();

            // 글꼴 크기
            result.FontSize = reader.ReadHUnit();

            // 언어 별 글꼴 인덱스
            result.FontIndicesByLanguage = reader.ReadBytes(7);

            // 언어 별 장평 비율
            result.SetWidthByLanguage = reader.ReadBytes(7);

            // 언어 별 자간 비율
            result.CharacterWidthByLanguage = reader.ReadSBytes(7);

            // 음영색 인덱스
            result.ShadeIndex = reader.ReadByte();

            // 글자색 인덱스
            result.TextColorIndex = reader.ReadByte();

            // 음영 비율
            result.ShadeRate = reader.ReadByte();

            // 속성
            result.Property = reader.ReadByte();

            // 예약 필드
            result.Reserved = reader.ReadBytes(4);

            return result;
        }

        public ushort FontSize { get; set; }

        public byte[] FontIndicesByLanguage { get; set; }

        public byte[] SetWidthByLanguage { get; set; }

        public sbyte[] CharacterWidthByLanguage { get; set; }

        public byte ShadeIndex { get; set; }

        public byte TextColorIndex { get; set; }

        public byte ShadeRate { get; set; }

        public byte Property { get; set; }

        public byte[] Reserved { get; set; }
    }

    public sealed class HwpV3ParagraphShape
    {
        public static HwpV3ParagraphShape Create(BinaryReader reader)
        {
            HwpV3ParagraphShape result = new HwpV3ParagraphShape();

            // 문단 왼쪽 여백
            result.LeftMargin = reader.ReadHUnit();

            // 문단 오른쪽 여백
            result.RightMargin = reader.ReadHUnit();

            // 문단 들여쓰기
            result.Indent = reader.ReadSHUnit();

            // 문단 줄 간격
            result.LineSpacing = reader.ReadHUnit();

            // 문안 아래 간격
            result.BottomMargin = reader.ReadHUnit();

            // 낱말 간격
            result.WordSpacing = reader.ReadByte();

            // 정렬 방식
            result.Alignment = reader.ReadByte();

            // 탭 설정
            result.TabSettings = HwpV3TabSettingsCollection.Create(reader);

            // 단 정의
            result.ColumnSettings = HwpV3ColumnSettings.Create(reader);

            // 음영 비율
            result.ShadeRate = reader.ReadByte();

            // 문단 테두리
            result.Border = reader.ReadByte();

            // 선 연결
            result.LineConnection = reader.ReadByte();

            // 문단 위 간격
            result.TopMargin = reader.ReadHUnit();

            // 예약
            result.Reserved = reader.ReadBytes(2);

            return result;
        }

        public ushort LeftMargin { get; set; }

        public ushort RightMargin { get; set; }

        public short Indent { get; set; }

        public ushort LineSpacing { get; set; }

        public ushort BottomMargin { get; set; }

        public byte WordSpacing { get; set; }

        public byte Alignment { get; set; }

        public HwpV3TabSettingsCollection TabSettings { get; set; }

        public HwpV3ColumnSettings ColumnSettings { get; set; }

        public byte ShadeRate { get; set; }

        public byte Border { get; set; }

        public byte LineConnection { get; set; }

        public ushort TopMargin { get; set; }

        public byte[] Reserved { get; set; }
    }

    public sealed class HwpV3TabSettings
    {
        public static HwpV3TabSettings Create(BinaryReader reader)
        {
            HwpV3TabSettings result = new HwpV3TabSettings();

            // 탭 종류
            result.TabType = reader.ReadByte();

            // 점 끌기 여부
            result.DotDecoration = reader.ReadByte();

            // 탭 위치
            result.TabPosition = reader.ReadHUnit();

            return result;
        }

        public byte TabType { get; set; }

        public byte DotDecoration { get; set; }

        public ushort TabPosition { get; set; }
    }

    public sealed class HwpV3TabSettingsCollection : List<HwpV3TabSettings>
    {
        public static HwpV3TabSettingsCollection Create(BinaryReader reader)
        {
            HwpV3TabSettingsCollection tabSettings = new HwpV3TabSettingsCollection();

            for (int i = 0; i < 40; i++)
                tabSettings.Add(HwpV3TabSettings.Create(reader));

            return tabSettings;
        }
    }

    public sealed class HwpV3AdditionalBlock
    {
        public static HwpV3AdditionalBlock Create(BinaryReader reader)
        {
            HwpV3AdditionalBlock result = new HwpV3AdditionalBlock();
            result.BlockType = 1;

            result.ID = reader.ReadDWord();
            result.Length = reader.ReadDWord();

            if (result.ID == 0u && result.Length == 0u)
                return result;

            result.Content = reader.ReadBytes((int)result.Length);
            return result;
        }

        public int BlockType { get; set; }

        public uint ID { get; set; }

        public uint Length { get; set; }

        public byte[] Content { get; set; }
    }

    public sealed class HwpV3AdditionalBlockCollection : List<HwpV3AdditionalBlock>
    {
        public static HwpV3AdditionalBlockCollection Create(BinaryReader reader)
        {
            HwpV3AdditionalBlockCollection results = new HwpV3AdditionalBlockCollection();

            while (true)
            {
                HwpV3AdditionalBlock block = HwpV3AdditionalBlock.Create(reader);
                if (block.ID == 0u && block.Length == 0u)
                    break;
                results.Add(block);
            }

            return results;
        }
    }

    public sealed class HwpV3Paragraph : List<HwpV3Paragraph>
    {
        public static HwpV3Paragraph Create(BinaryReader reader)
        {
            HwpV3Paragraph p = null;
            HwpV3Paragraph paragraph = new HwpV3Paragraph();
            while (true)
            {
                p = Create(reader, paragraph);
                if (p == null)
                    break;
                paragraph.Add(p);
            }
            return paragraph;
        }

        private static HwpV3Paragraph Create(BinaryReader reader, HwpV3Paragraph parent)
        {
            byte prev_paragraph_shape;
            ushort n_chars;
            ushort n_lines;
            byte char_shape_included;

            byte flag;
            int i;

            prev_paragraph_shape = reader.ReadByte();
            n_chars = reader.ReadUInt16();
            n_lines = reader.ReadUInt16();
            char_shape_included = reader.ReadByte();

            reader.ReadBytes(1 + 4 + 1 + 31);
            // 여기까지 43 바이트

            if (prev_paragraph_shape == 0 && n_chars > 0)
            {
                reader.ReadBytes(187);
            }

            // 빈 문단이면 null 반환
            if (n_chars == 0)
                return null;

            // 줄 정보
            reader.ReadBytes(n_lines * 14);

            // 글자 모양 정보
            if (char_shape_included != 0)
            {
                for (i = 0; i < n_chars; i++)
                {
                    flag = reader.ReadByte();

                    if (flag == 1)
                        continue;

                    reader.ReadBytes(31);
                }
            }

            HwpV3Paragraph p = null;
            HwpV3Paragraph paragraph = new HwpV3Paragraph();
            string @string = String.Empty;

            // 글자들
            ushort n_chars_read = 0;
            ushort c;

            while (n_chars_read < n_chars)
            {
                c = reader.ReadUInt16();
                n_chars_read += 1;

                if (c == 6)
                {
                    n_chars_read += 3;
                    reader.ReadBytes(6 + 34);
                    continue;
                }
                else if (c == 9)
                { /* tab */
                    n_chars_read += 3;
                    reader.ReadBytes(6);
                    @string += "\t";
                    continue;
                }
                else if (c == 10)
                { /* table */
                    n_chars_read += 3;
                    reader.ReadBytes(6);

                    /* 테이블 식별 정보 84 바이트 */
                    reader.ReadBytes(80);

                    ushort n_cells = reader.ReadUInt16();

                    reader.ReadBytes(2);
                    reader.ReadBytes(27 * n_cells);

                    /* <셀 문단 리스트>+ */
                    for (i = 0; i < n_cells; i++)
                    {
                        /* <셀 문단 리스트> ::= <셀 문단>+ <빈문단> */
                        while (true)
                        {
                            p = Create(reader, parent);
                            if (p == null)
                                break;
                            paragraph.Add(p);
                        }
                    }

                    /* <캡션 문단 리스트> ::= <캡션 문단>+ <빈문단> */
                    while (true)
                    {
                        p = Create(reader, parent);
                        if (p == null)
                            break;
                        paragraph.Add(p);
                    }
                    continue;
                }
                else if (c == 11)
                {
                    n_chars_read += 3;

                    reader.ReadBytes(6);
                    uint len2 = reader.ReadUInt32();
                    reader.ReadBytes(344);
                    reader.ReadBytes((int)len2);

                    /* <캡션 문단 리스트> ::= <캡션 문단>+ <빈문단> */
                    while (true)
                    {
                        p = Create(reader, parent);
                        if (p == null)
                            break;
                        paragraph.Add(p);
                    }
                    continue;
                }
                else if (c == 13)
                { /* 글자들 끝 */
                    @string += Environment.NewLine;
                    continue;
                }
                else if (c == 16)
                {
                    n_chars_read += 3;
                    reader.ReadBytes(6);
                    reader.ReadBytes(10);
                    /* <문단 리스트> ::= <문단>+ <빈문단> */
                    while (true)
                    {
                        p = Create(reader, parent);
                        if (p == null)
                            break;
                        paragraph.Add(p);
                    }
                    continue;
                }
                else if (c == 17)
                { /* 각주/미주 */
                    n_chars_read += 3;
                    reader.ReadBytes(6);
                    reader.ReadBytes(14);
                    while (true)
                    {
                        p = Create(reader, parent);
                        if (p == null)
                            break;
                        paragraph.Add(p);
                    }
                    continue;
                }
                else if (c == 18 || c == 19 || c == 20 || c == 21)
                {
                    n_chars_read += 3;
                    reader.ReadBytes(6);
                    continue;
                }
                else if (c == 23)
                { /*글자 겹침 */
                    n_chars_read += 4;
                    reader.ReadBytes(8);
                    continue;
                }
                else if (c == 24 || c == 25)
                {
                    n_chars_read += 2;
                    reader.ReadBytes(4);
                    continue;
                }
                else if (c == 28)
                { /* 개요 모양/번호 */
                    n_chars_read += 31;
                    reader.ReadBytes(62);
                    continue;
                }
                else if (c == 30 || c == 31)
                {
                    n_chars_read += 1;
                    reader.ReadBytes(2);
                    continue;
                }
                else if (c >= 0x0020 && c <= 0xffff)
                {
                    string tmp = new String(Hnc2Unicode.hnchar_to_utf8(c));
                    @string += tmp;
                    continue;
                }
                else
                {
#if DEBUG
                    Trace.WriteLine("Special Character: {0}", ((int)c).ToString("X4"));
#endif
                } /* if */
            }

            paragraph.Text = @string;
            return paragraph;
        }

        public string Text { get; set; }

        public override string ToString()
        { return this.ToString(true); }

        public string ToString(bool includeChildParagraphs)
        {
            StringBuilder buffer = new StringBuilder(this.Text);

            if (includeChildParagraphs)
                foreach (HwpV3Paragraph each in this)
                    buffer.AppendLine(each.ToString(includeChildParagraphs));

            return buffer.ToString();
        }
    }
}
