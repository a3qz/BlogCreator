// Learn more about F# at http://fsharp.org
// See the 'F# Tutorial' project for more help.

open System.IO
open System.Diagnostics
open System.IO.Packaging
open DocumentFormat.OpenXml.Packaging 
open DocumentFormat.OpenXml.Wordprocessing 
open System.Text
open DocumentFormat.OpenXml
open System

let _PATH = "C:\\Users\\Ryan\\Documents\\ethicsblogs\\"
let Main args = 

    Directory.GetFiles("C:\Users\Ryan\Documents\ethicsblogs\test", "*") 
    |> Array.map Path.GetFileName 
    |> Array.iter (printfn "%s") 

let xmlSeq = Seq.initInfinite (fun index -> sprintf "<author><name>name%d</name><age>%d</age><books><book>book%d</book></books></author>" index index index)

let createFile (seq: string seq) numberToTake fileName =
    use streamWriter = new StreamWriter("C:\\Users\\Ryan\\Documents\\" + fileName, false)
    streamWriter.WriteLine("<startTag>")
    let rec internalWriter (seq: string seq) (sw:StreamWriter) i (endTag:string) =
        match i with
        | 0 -> (sw.WriteLine(Seq.head seq);
            sw.WriteLine(endTag))
        | _ -> (sw.WriteLine(Seq.head seq);
            internalWriter (Seq.skip 1 seq) sw (i-1) endTag)
    internalWriter seq streamWriter numberToTake "</startTag>"

let funcTimer fn =
    let stopWatch = Stopwatch.StartNew()
    printfn "Timing started"
    fn()
    stopWatch.Stop()
    printfn "Time elased: %A" stopWatch.Elapsed

        
let rec GetPlainText (element : OpenXmlElement ) : string= 
    let PlainTextInWord = new StringBuilder()
    let x = 
        element.Elements() 
        |> Seq.iter (fun (v: OpenXmlElement) -> match v.LocalName with 
                                                | "t" -> PlainTextInWord.Append v.InnerText |> ignore
                                                | "cr" -> PlainTextInWord.Append Environment.NewLine|> ignore
                                                | "br" -> PlainTextInWord.Append Environment.NewLine|> ignore
                                                | "tab" -> PlainTextInWord.Append "\t"|> ignore
                                                | "p" ->
                                                    PlainTextInWord.Append (GetPlainText v) |> ignore
                                                    PlainTextInWord.AppendLine Environment.NewLine |> ignore
                                                | _ -> PlainTextInWord.Append (GetPlainText v)|> ignore)
        
 
    string PlainTextInWord

let fileReader filename = 
    let header = _PATH
    let fullname = header + filename
    let fullname2 = header + "2"
    let fullname2 = fullname2 + filename
    let byteArray = File.ReadAllBytes fullname
    
    use mem = new MemoryStream() 
    mem.Write(byteArray, 0, (int)byteArray.Length) 

    do
        use doc = WordprocessingDocument.Open(mem, true) 
        let sb = new StringBuilder();
        let bod = doc.MainDocumentPart.Document.Body; 
        let plaintext = GetPlainText bod
        printfn "%s" plaintext

    use fs = new FileStream(fullname2, FileMode.Create) 
    mem.WriteTo(fs) 


(funcTimer (fun () -> createFile xmlSeq 100 "file100.xml"))
Directory.GetFiles(_PATH, "*") 
    |> Array.map Path.GetFileName 
    |> Array.iter fileReader 
    |> ignore
System.Console.ReadLine();
