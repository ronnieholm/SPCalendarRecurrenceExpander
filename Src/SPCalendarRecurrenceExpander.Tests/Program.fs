module SPCalendarRecurrenceExpander.Tests.Program

open System
open System.Reflection
open NUnit.Core
open NUnit.Framework
open Swensen.Unquote

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