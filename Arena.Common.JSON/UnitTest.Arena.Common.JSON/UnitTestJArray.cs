// UnitTestJArray.cs - 06/14/2017

using Arena.Common.JSON;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTest.Arena.Common.JSON
{
    [TestClass]
    public class UnitTestJArray
    {
        [TestMethod]
        public void TestNullJArrayWithWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[]";
            JArray ja = new JArray();
            // act
            actualValue = ja.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestNullJArrayNoWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[]";
            JArray ja = new JArray();
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestNullJArrayDefaultWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[]";
            JArray ja = new JArray();
            // act
            actualValue = ja.ToString();
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayNullValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[null]";
            JArray ja = new JArray();
            ja.Add(null);
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayNullValueWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\r\n  null\r\n]";
            JArray ja = new JArray();
            ja.Add(null);
            // act
            actualValue = ja.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayFalseValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[false]";
            JArray ja = new JArray();
            ja.Add(false);
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayTrueValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[true]";
            JArray ja = new JArray();
            ja.Add(true);
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayStringValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\"abc\"]";
            JArray ja = new JArray();
            ja.Add("abc");
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayIntValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[123]";
            JArray ja = new JArray();
            ja.Add(123);
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayDoubleValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[123.45]";
            JArray ja = new JArray();
            ja.Add(123.45);
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayJObjectValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[123.45,{\"key\":\"value\"}]";
            JArray ja = new JArray();
            JObject jo = new JObject();
            jo.Add("key", "value");
            ja.Add(123.45);
            ja.Add(jo);
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayJArrayValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[123.45,[\"key\",\"value\"]]";
            JArray ja = new JArray();
            JArray ja2 = new JArray();
            ja.Add(123.45);
            ja2.Add("key");
            ja2.Add("value");
            ja.Add(ja2);
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayMultiValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\"abc\",123.45]";
            JArray ja = new JArray();
            ja.Add("abc");
            ja.Add(123.45);
            // act
            actualValue = ja.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayMultiValueWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\r\n  \"abc\",\r\n  123.45,\r\n  null\r\n]";
            JArray ja = new JArray();
            ja.Add("abc");
            ja.Add(123.45);
            ja.Add(null);
            // act
            actualValue = ja.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayNewEmpty()
        {
            // arrange
            string actualValue;
            string expectedValue = "[]";
            JArray ja = JArray.Parse(expectedValue);
            // act
            actualValue = ja.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJArrayNewValues()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\r\n  \"abc\",\r\n  123.45,\r\n  null\r\n]";
            JArray ja = JArray.Parse(expectedValue);
            // act
            actualValue = ja.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

    }
}
