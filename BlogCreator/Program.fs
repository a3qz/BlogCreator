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

// Path to the folder where the blogs are stored
let BLOG_PATH = "C:\\Users\\Ryan\\Documents\\ethicsblogs\\"

// Create a file based on a sequence of strings and a filename
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


// Given an xml element, recursivley traverse it until it finds a value that can be converted to a string
let rec GetPlainText (element : OpenXmlElement ) : string= 
    let PlainTextInWord = new StringBuilder()
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

// Given a filename, return the text of that file as a string
let FileReader filename = 
    let header = BLOG_PATH
    let fullname = header + filename
    let byteArray = File.ReadAllBytes fullname
    
    use mem = new MemoryStream() 
    mem.Write(byteArray, 0, (int)byteArray.Length) 
    
    use doc = WordprocessingDocument.Open(mem, true) 
    let bod = doc.MainDocumentPart.Document.Body; 
    let plaintext = GetPlainText bod
    plaintext
    
// Effectivley main; get all the files in the directory and call FileReader on them
Directory.GetFiles(BLOG_PATH, "*") 
    |> Array.map Path.GetFileName 
    |> Array.map FileReader 
    |> Array.iter (printfn "%s")

// Makes it so that I can read the output without the window closing
System.Console.ReadLine();
