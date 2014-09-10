module SPCalendarRecurrenceExpander.Tests.Program

// <package id="NUnit.Runners" version="2.6.3" />
// add unquote package

// add references to
//   nunit.core
//   nunit.core.interfaces
//   nunit.framework

open System
open System.Reflection
open NUnit.Core
open NUnit.Framework
open Swensen.Unquote

// DOESN*T MATTER WHICH VERSION OF VS YOU USE OR IF YOU DON*T USE VS AT ALL
// DOESN*T PRECLUDE YOU FROM USING THE STANDARD NUNIT TEST RUNNERS

// alternatively, place tests in modules or types
//let [<Test>] ``successful nunit assertion``() = Assert.IsTrue(true)    
//let [<Test>] ``failing nunit assertion``() = Assert.IsTrue(false)    
//let [<Test>] ``successful unquote test``() = test <@ 42 = 42 @>        
//let [<Test>] ``failing unquote test``() = test <@ [1;2;3] = [4;5] @>

let eventListener() =
    { new EventListener with        
        member __.RunStarted(name: string, testCount: int) = printfn "Run started"
        member __.TestStarted(tn: TestName) = printf "%s" tn.FullName
        member __.TestOutput(t: TestOutput) = ()
        member __.TestFinished(tr: TestResult) = 
            match tr.ResultState with
            | ResultState.Failure -> printfn " => %O\r\n%s" tr.ResultState tr.Message
            | _ -> printfn " => %O" tr.ResultState
        member __.RunFinished(tr: TestResult) = printfn "Overall result: %O" tr.ResultState
        member __.RunFinished(e: Exception) = printfn "Exception:\r\n%s" e.StackTrace 
        member __.UnhandledException(e: Exception) = printfn "Exception:\r\n%s" e.StackTrace                
        member __.SuiteFinished(tr: TestResult) = ()
        member __.SuiteStarted(tn: TestName) = () }

[<EntryPoint>]
let main _ =
    CoreExtensions.Host.InitializeService()    
    let p = TestPackage("Tests")
    p.Assemblies.Add(Assembly.GetExecutingAssembly().Location) |> ignore
    use r = new SimpleTestRunner()
    r.Load(p) |> ignore
    r.Run(eventListener(), TestFilter.Empty, true, LoggingThreshold.All) |> ignore
    0