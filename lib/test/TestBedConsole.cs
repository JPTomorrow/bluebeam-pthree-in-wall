
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using C = System.Console;


namespace JPMorrow.Test.Console
{
    public class AssertInformation
    {
        public string Message { get; private set; } = "";
        public bool Result { get; private set; } = false;

        public AssertInformation(string msg, bool result)
        {
            Message = msg;
            Result = result;
        }
    }

    public class TestAssert 
    {
        private List<AssertInformation> asserts { get; set; } = new List<AssertInformation>();
        private Exception Exception { get; set; } = null;
        public bool HasPassed { get => asserts.All(x => x.Result == true); }

        public TestAssert() {}

        public void Assert(string message, bool condition)
        {
            asserts.Add(new AssertInformation(message, condition));
        }

        public void AssignFailException(Exception e)
        {
            Exception = e;
        }

        public string GetFailMessage()
        {
            if(HasPassed) return "";

            if(Exception == null)
            {
                var idx = asserts.FindIndex(x => x.Result == false);
                if(idx == -1) return "";
                else
                {
                    return asserts[idx].Message;
                }
            }

            return Exception.Message;
        }
    }

    public class TestResult
    {
        public string FunctionName { get; private set; }
        public TestAssert Assert { get; private set; }
        public string Result { get => Assert.HasPassed ? "Passed" : "Failed"; }

        private ConsoleColor passedFailColor { get {
            if(Result.Equals("Passed")) return ConsoleColor.Green;
            else if(Result.Equals("Failed")) return ConsoleColor.Red;
            else return ConsoleColor.White;
            }}
        
        // stops inlining so we can find the stacktrace function name reliably
        [MethodImpl(MethodImplOptions.NoInlining)] 
        public TestResult(string function_name, TestAssert assert)
        {
            string func_name = function_name; // new StackFrame(1, true).GetMethod().Name;
            func_name = Regex.Replace(func_name, "(\\B[A-Z])", " $1");
            FunctionName = func_name;
            Assert = assert;
        }
        
        /// <summary>
        /// Print all the provided Test Results
        /// </summary>
        /// <param name="results"></param>
        public static void PrintAllTests(IEnumerable<TestResult> results)
        {
            C.WriteLine();
            C.WriteLine("Test Results:");

            foreach(var r in results)
            {
                C.Write(r.FunctionName + " -> ");

                if(!r.Assert.HasPassed)
                {
                    var exception_msg = r.Assert.GetFailMessage();
                    if(!string.IsNullOrWhiteSpace(exception_msg)) C.Write(exception_msg + " -> ");
                }

                TestBedConsoleUtils.ConsoleColorWriteLine(r.Result, r.passedFailColor);
            }
        }
    }
    
    public static partial class TestBedConsole
    {
        public static void TestAll(string exe_path)
        {
            var methods = typeof(TestBed).GetMethods()
                .Where(x => !x.Name.Contains("TestAll") && x.Name.ToLower().StartsWith("test"));

            List<TestResult> results = new List<TestResult>();

            TestResult entry_failed_on = null;

            foreach(var m in methods)
            {
                var t = new TestAssert();

                try
                {
                    m.Invoke(null, new object[] { exe_path, t });
                }
                catch(Exception ex)
                {
                    t.AssignFailException(ex);
                }

                var result = new TestResult(m.Name, t);
                results.Add(result);
                
                if(!result.Assert.HasPassed)
                {
                    entry_failed_on = result;
                    break;
                }
            }

            TestResult.PrintAllTests(results);
            var finished_txt = "\nTests Finished:\n";

            if(entry_failed_on != null)
            {
                C.WriteLine(finished_txt);
                C.Write("\tResult: ");
                TestBedConsoleUtils.ConsoleColorWrite("Failed", ConsoleColor.Red);
                C.WriteLine();
                C.WriteLine("\tFailed Test Name: " + entry_failed_on.FunctionName);
                C.WriteLine("\tReason: " + entry_failed_on.Assert.GetFailMessage());
            }
            else
            {
                C.WriteLine(finished_txt);
                C.Write("\tResult: ");
                TestBedConsoleUtils.ConsoleColorWrite("Passed", ConsoleColor.Green);
            }
        }
    }

    public static class TestBedConsoleUtils
    {
        public static void ConsoleColorWrite(string txt, ConsoleColor color)
        {
            C.ForegroundColor = color;
            C.Write(txt);
            C.ResetColor();
        }

        public static void ConsoleColorWriteLine(string txt, ConsoleColor color)
        {
            C.ForegroundColor = color;
            C.WriteLine(txt);
            C.ResetColor();
        }
    }
}