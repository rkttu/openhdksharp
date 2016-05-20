using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClouDeveloper.HWP.Test
{
    [TestFixture]
    public class CharacterSetTestFixture
    {
        [Test]
        public void InitializerTest()
        {
            Assert.That(
                CharacterSet.hnc2uni_page0,
                Is.Not.Null.And.Length.EqualTo(0x3fff - 0x0000 + 1));
            Assert.That(
                CharacterSet.hnc2uni_page4,
                Is.Not.Null.And.Length.EqualTo(0x5317 - 0x4000 + 1));
            Assert.That(
                CharacterSet.hnc2uni_page5,
                Is.Not.Null.And.Length.EqualTo(0x7fff - 0x5318 + 1));
            Assert.That(
                CharacterSet.hnc2uni_page8,
                Is.Not.Null.And.Length.EqualTo((0xffff - 0x8000 + 1) * 3));
            Assert.That(
                CharacterSet.hyc2uni_page14,
                Is.Not.Null.And.Length.EqualTo((0xF8F7 - 0xE0BC + 1) * 3));

            Assert.That(
                CharacterSet.hnc2uni_page0,
                Is.EqualTo(CharacterSet._hnc2uni_page0).Within(0));
            Assert.That(
                CharacterSet.hnc2uni_page4,
                Is.EqualTo(CharacterSet._hnc2uni_page4).Within(0));
            Assert.That(
                CharacterSet.hnc2uni_page5,
                Is.EqualTo(CharacterSet._hnc2uni_page5).Within(0));
            Assert.That(
                CharacterSet.hnc2uni_page8,
                Is.EqualTo(CharacterSet._hnc2uni_page8).Within(0));
            Assert.That(
                CharacterSet.hyc2uni_page14,
                Is.EqualTo(CharacterSet._hyc2uni_page14).Within(0));
        }
    }
}
