using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelHelper.Tests
{
    [TestClass()]
    public class ExcelHelperTests
    {
        public class TestClass
        {
            /// <summary>
            /// 第一項
            /// </summary>
            public string one
            {
                get;set;
            }


            public string two
            {
                get;set;
            } 

            public string three
            {
                get;set;
            }

            public string four
            {
                get;set;
            }

            public string six
            {
                get;set;
            }

            public string seven
            {
                get;set;
            }
            
        }

        [TestMethod()]
        public void ImportExcelAsyncTest()
        {
            //arrange

            string path = @"D:\Coding\WorkProject\Githubs\DerekTools\ExcelTools\ExcelHelperTests\TestFile\一般.xlsx";
            var fileInfo = new FileInfo(path);
            var fileByte = File.ReadAllBytes(path);

            //action
            var actual = fileByte.ImportExcelAsync<TestClass>();
            //assert
            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.Count()==1);
            Assert.IsNotNull(actual.First().one);
        }

        [TestMethod()]
        public void ImportExcelAsync_Define_Test()
        {
            //arrange

            string path = @"D:\Coding\WorkProject\Githubs\DerekTools\ExcelTools\ExcelHelperTests\TestFile\定義.xlsx";
            var fileInfo = new FileInfo(path);
            var fileByte = File.ReadAllBytes(path);

            //action
            var actual = fileByte.ImportExcelAsync<TestClass>();
            //assert
            Assert.IsNotNull(actual);
            Assert.IsTrue(actual.Count() == 1);
            Assert.IsNotNull(actual.First().one);
        }
    }
}