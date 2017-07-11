// UnitTestJObject.cs - 06/14/2017

using Arena.Common.JSON;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace UnitTest.Arena.Common.JSON
{
    [TestClass]
    public class UnitTestJObject
    {
        [TestMethod]
        public void TestNullJObjectWithWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{}";
            JObject jo = new JObject();
            // act
            actualValue = jo.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestNullJObjectNoWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{}";
            JObject jo = new JObject();
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestNullJObjectDefaultWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{}";
            JObject jo = new JObject();
            // act
            actualValue = jo.ToString();
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectNullValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":null}";
            JObject jo = new JObject();
            jo.Add("key", null);
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectFalseValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":false}";
            JObject jo = new JObject();
            jo.Add("key", false);
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectTrueValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":true}";
            JObject jo = new JObject();
            jo.Add("key", true);
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectStringValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":\"abc\"}";
            JObject jo = new JObject();
            jo.Add("key", "abc");
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectStringValueCtrlChars()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":\"\\r\\n\\t\\b\\f\\u1234\"}";
            JObject jo = new JObject();
            jo.Add("key", "\r\n\t\b\f\u1234");
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectIntValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":123}";
            JObject jo = new JObject();
            jo.Add("key", 123);
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectIntValueWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\r\n  \"key\": 123\r\n}";
            JObject jo = new JObject();
            jo.Add("key", 123);
            // act
            actualValue = jo.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectDoubleValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":123.45}";
            JObject jo = new JObject();
            jo.Add("key", 123.45);
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectDoubleExponentValue()
        {
            // arrange
            double actualValue;
            double expectedValue = 1.2345e50;
            JObject jo = JObject.Parse("{\"key\":1.2345e50}");
            // act
            actualValue = (double)jo.GetValue("key");
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectDoubleExponentNegativeValue()
        {
            // arrange
            double actualValue;
            double expectedValue = 1.2345e-50;
            JObject jo = JObject.Parse("{\"key\":1.2345e-50}");
            // act
            actualValue = (double)jo.GetValue("key");
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectDoubleExponentUpperValue()
        {
            // arrange
            double actualValue;
            double expectedValue = 1.2345E50;
            JObject jo = JObject.Parse("{\"key\":1.2345E50}");
            // act
            actualValue = (double)jo.GetValue("key");
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectDoubleExponentUpperNegativeValue()
        {
            // arrange
            double actualValue;
            double expectedValue = 1.2345E-50;
            JObject jo = JObject.Parse("{\"key\":1.2345E-50}");
            // act
            actualValue = (double)jo.GetValue("key");
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectDateValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":\"2017-01-02T00:00:00.0000000\"}";
            JObject jo = new JObject();
            jo.Add("key", DateTime.Parse("01/02/2017"));
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectDatetimeValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":\"2017-01-02T16:42:25.0000000\"}";
            JObject jo = new JObject();
            jo.Add("key", DateTime.Parse("01/02/2017 16:42:25"));
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectJObjectValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":{\"newkey\":456}}";
            JObject jo = new JObject();
            JObject jo2 = new JObject();
            jo2.Add("newkey", 456);
            jo.Add("key", jo2);
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectJArrayValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":[\"newkey\",456]}";
            JObject jo = new JObject();
            JArray ja = new JArray();
            ja.Add("newkey");
            ja.Add(456);
            jo.Add("key", ja);
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectMultiValue()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\"key\":123,\"otherkey\":789.12}";
            JObject jo = new JObject();
            jo.Add("key", 123);
            jo.Add("otherkey", 789.12);
            // act
            actualValue = jo.ToString(JsonFormat.None);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectMultiValueWhitespace()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\r\n  \"key\": 123,\r\n  \"otherkey\": 789.12\r\n}";
            JObject jo = new JObject();
            jo.Add("key", 123);
            jo.Add("otherkey", 789.12);
            // act
            actualValue = jo.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectNewEmpty()
        {
            // arrange
            string actualValue;
            string expectedValue = "{}";
            JObject jo = JObject.Parse(expectedValue);
            // act
            actualValue = jo.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        [TestMethod]
        public void TestJObjectNewValues()
        {
            // arrange
            string actualValue;
            string expectedValue = "{\r\n  \"key\": 123,\r\n  \"otherkey\": 789.12\r\n}";
            JObject jo = JObject.Parse(expectedValue);
            // act
            actualValue = jo.ToString(JsonFormat.Indent);
            // assert
            Assert.AreEqual(expectedValue, actualValue);
        }
    }
}
