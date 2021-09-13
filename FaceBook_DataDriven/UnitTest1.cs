using System;
using System.Diagnostics;
using System.IO;
using ExcelDataReader;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace FaceBook_DataDriven
{
    
    public class ReadDataFromExcel:BaseClass
    {
       

        [Test]

        public void ReadingData()
        {
            
            ExcelOperation.PopulateInCollection(@"C:\Users\HP\Documents\data.xlsx");
            Debug.WriteLine("******");
            driver.FindElement(By.Name("email")).SendKeys(ExcelOperation.ReadData(1, "email"));
            System.Threading.Thread.Sleep(2000);
            driver.FindElement(By.Id("pass")).SendKeys(ExcelOperation.ReadData(1, "password"));
            System.Threading.Thread.Sleep(2000);
            driver.FindElement(By.Name("login")).Click();
            System.Threading.Thread.Sleep(12000);



        }
    }
}