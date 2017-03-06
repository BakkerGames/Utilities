using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Common.JSON;

namespace UnitTest.Common.JSON
{
    [TestClass]
    public class UnitTestJSONObject
    {
        [TestMethod]
        public void TestNullJSONObjectWithWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{}";
            JSONObject jo = new JSONObject();
            // act
            actualValue = jo.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestNullJSONObjectNoWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{}";
            JSONObject jo = new JSONObject();
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestNullJSONObjectDefaultWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{}";
            JSONObject jo = new JSONObject();
            // act
            actualValue = jo.ToString();
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectNullValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":null}";
            JSONObject jo = new JSONObject();
            jo.Add("key", null);
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectFalseValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":false}";
            JSONObject jo = new JSONObject();
            jo.Add("key", false);
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectTrueValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":true}";
            JSONObject jo = new JSONObject();
            jo.Add("key", true);
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectStringValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":\"abc\"}";
            JSONObject jo = new JSONObject();
            jo.Add("key", "abc");
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectStringValueCtrlChars()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":\"\\r\\n\\t\\b\\f\\u1234\"}";
            JSONObject jo = new JSONObject();
            jo.Add("key", "\r\n\t\b\f\u1234");
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectIntValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":123}";
            JSONObject jo = new JSONObject();
            jo.Add("key", 123);
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectIntValueWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\r\n  \"key\": 123\r\n}";
            JSONObject jo = new JSONObject();
            jo.Add("key", 123);
            // act
            actualValue = jo.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectDoubleValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":123.45}";
            JSONObject jo = new JSONObject();
            jo.Add("key", 123.45);
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectDateValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":\"2017-01-02T00:00:00.0000000\"}";
            JSONObject jo = new JSONObject();
            jo.Add("key", DateTime.Parse("01/02/2017"));
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectDatetimeValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":\"2017-01-02T16:42:25.0000000\"}";
            JSONObject jo = new JSONObject();
            jo.Add("key", DateTime.Parse("01/02/2017 16:42:25"));
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectJSONObjectValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":{\"newkey\":456}}";
            JSONObject jo = new JSONObject();
            JSONObject jo2 = new JSONObject();
            jo2.Add("newkey", 456);
            jo.Add("key", jo2);
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectJSONArrayValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":[\"newkey\",456]}";
            JSONObject jo = new JSONObject();
            JSONArray ja = new JSONArray();
            ja.Add("newkey");
            ja.Add(456);
            jo.Add("key", ja);
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectMultiValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":123,\"otherkey\":789.12}";
            JSONObject jo = new JSONObject();
            jo.Add("key", 123);
            jo.Add("otherkey", 789.12);
            // act
            actualValue = jo.ToString(false);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectMultiValueWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\r\n  \"key\": 123,\r\n  \"otherkey\": 789.12\r\n}";
            JSONObject jo = new JSONObject();
            jo.Add("key", 123);
            jo.Add("otherkey", 789.12);
            // act
            actualValue = jo.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectNewEmpty()
        {
            // arrange
            string actualValue;
            string expectedValue = "{}";
            JSONObject jo = new JSONObject(expectedValue);
            // act
            actualValue = jo.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJSONObjectNewValues()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\r\n  \"key\": 123,\r\n  \"otherkey\": 789.12\r\n}";
            JSONObject jo = new JSONObject(expectedValue);
            // act
            actualValue = jo.ToString(true);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

    }
}
