using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Common.JSON;

namespace UnitTest.Common.JSON
{
    [TestClass]
    public class UnitTestJSONArray
    {
        [TestMethod]
        public void TestNullJSONArrayWithWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[]";
            JSONArray ja = new JSONArray();
            // act
            actualValue = ja.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestNullJSONArrayNoWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[]";
            JSONArray ja = new JSONArray();
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestNullJSONArrayDefaultWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[]";
            JSONArray ja = new JSONArray();
            // act
            actualValue = ja.ToString();
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayNullValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[null]";
            JSONArray ja = new JSONArray();
            ja.Add(null);
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayNullValueWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\r\n  null\r\n]";
            JSONArray ja = new JSONArray();
            ja.Add(null);
            // act
            actualValue = ja.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayFalseValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[false]";
            JSONArray ja = new JSONArray();
            ja.Add(false);
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayTrueValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[true]";
            JSONArray ja = new JSONArray();
            ja.Add(true);
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayStringValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\"abc\"]";
            JSONArray ja = new JSONArray();
            ja.Add("abc");
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayIntValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[123]";
            JSONArray ja = new JSONArray();
            ja.Add(123);
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayDoubleValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[123.45]";
            JSONArray ja = new JSONArray();
            ja.Add(123.45);
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayJSONObjectValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[123.45,{\"key\":\"value\"}]";
            JSONArray ja = new JSONArray();
            JSONObject jo = new JSONObject();
            jo.Add("key", "value");
            ja.Add(123.45);
            ja.Add(jo);
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayJSONArrayValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[123.45,[\"key\",\"value\"]]";
            JSONArray ja = new JSONArray();
            JSONArray ja2 = new JSONArray();
            ja.Add(123.45);
            ja2.Add("key");
            ja2.Add("value");
            ja.Add(ja2);
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayMultiValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\"abc\",123.45]";
            JSONArray ja = new JSONArray();
            ja.Add("abc");
            ja.Add(123.45);
            // act
            actualValue = ja.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayMultiValueWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\r\n  \"abc\",\r\n  123.45,\r\n  null\r\n]";
            JSONArray ja = new JSONArray();
            ja.Add("abc");
            ja.Add(123.45);
            ja.Add(null);
            // act
            actualValue = ja.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayNewEmpty()
        {
            // arrange
            string actualValue;
            string expectedValue = "[]";
            JSONArray ja = new JSONArray(expectedValue);
            // act
            actualValue = ja.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONArrayNewValues()
        {
            // arrange
            string actualValue;
            string expectedValue = "[\r\n  \"abc\",\r\n  123.45,\r\n  null\r\n]";
            JSONArray ja = new JSONArray(expectedValue);
            // act
            actualValue = ja.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

    }
}
