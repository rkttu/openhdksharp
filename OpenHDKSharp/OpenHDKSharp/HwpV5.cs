/*
 * HwpV5.cs - Under Construction
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
using System.IO;
using System.Linq;
using System.Text;

namespace OpenHDKSharp.V5
{
    public static class HwpV5BinaryReaderExtension
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

    public sealed class HwpV5Document
    {
    }
}
