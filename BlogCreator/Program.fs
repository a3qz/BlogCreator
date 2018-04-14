// Learn more about F# at http://fsharp.org
// See the 'F# Tutorial' project for more help.
open FSharp.Data
open System.IO
open System.Diagnostics
open System.IO.Packaging
open System.Web.UI
open DocumentFormat.OpenXml.Packaging 
open DocumentFormat.OpenXml.Wordprocessing 
open System.Text
open DocumentFormat.OpenXml
open System

// Path to the folder where the blogs are stored

System.Environment.GetCommandLineArgs() |> printfn "args: %A" 
System.Environment.GetCommandLineArgs().[1] |> printfn "args: %s"
System.Console.ReadLine() |> ignore

let BLOG_CLASS_HARD = "history"
let BLOG_CLASS =  if System.Environment.GetCommandLineArgs().Length > 1 then System.Environment.GetCommandLineArgs().[1] else BLOG_CLASS_HARD


let BLOG_DOCUMENT_SOURCE_PATH = "C:\\Users\\Ryan\\Documents\\"+BLOG_CLASS+"blogs\\"
let BLOG_HTML_DESTINTION_PATH = "C:\\Users\\Ryan\\Documents\\a3qz.github.io\\"+BLOG_CLASS+"\\blogs\\"
let SITE_RSS_DESTINATION_PATH = "C:\\Users\\Ryan\\Documents\\a3qz.github.io\\"+BLOG_CLASS+"\\rss.xml"

// Given an xml element, recursivley traverse it until it finds a value that can be converted to a string
let rec GetPlainText (element : OpenXmlElement ) : string= 
    let PlainTextInWord = new StringBuilder()
    element.Elements() 
        |> Seq.iter (fun (v: OpenXmlElement) -> match v.LocalName with 
                                                | "t" -> PlainTextInWord.Append v.InnerText |> ignore
                                                | "cr" -> PlainTextInWord.Append "<br />"|> ignore
                                                | "br" -> PlainTextInWord.Append "<br />"|> ignore
                                                | "tab" -> PlainTextInWord.Append "&emsp;"|> ignore
                                                | "p" ->
                                                    PlainTextInWord.Append (GetPlainText v) |> ignore
                                                    PlainTextInWord.AppendLine "<br />" |> ignore
                                                | _ -> PlainTextInWord.Append (GetPlainText v)|> ignore)
        
 
    string PlainTextInWord

// Given a filename, return the text of that file as a string
let FileReader filename = 
    let fullname = BLOG_DOCUMENT_SOURCE_PATH + filename
    let byteArray = File.ReadAllBytes fullname
    
    use mem = new MemoryStream() 
    mem.Write(byteArray, 0, (int)byteArray.Length) 
    
    use doc = WordprocessingDocument.Open(mem, true) 
    let bod = doc.MainDocumentPart.Document.Body; 
    let plaintext = GetPlainText bod
    plaintext

// Taken from stackoverflow
let (|Prefix|_|) (p:string) (s:string) =
    if s.StartsWith(p) then
        Some(s.Substring(p.Length))
    else
        None

// Take a string of html with my stupid ~ as delimiter thing and split it, then write out the html 
let HTMLMaker (text: string) = 
    let stringWriter = new StringWriter()
    use writer = new HtmlTextWriter(stringWriter)

    text.Split [|'~'|]
    |> Seq.iter (fun t -> match t with 
                                | Prefix "Reading" rest -> 
                                                writer.RenderBeginTag HtmlTextWriterTag.Html
                                                writer.RenderBeginTag HtmlTextWriterTag.Head
                                                writer.RenderBeginTag HtmlTextWriterTag.Title
                                                writer.Write t
                                                writer.RenderEndTag()
                                                writer.RenderEndTag()
                                                writer.RenderBeginTag HtmlTextWriterTag.Body
                                                writer.RenderBeginTag HtmlTextWriterTag.H1
                                                writer.Write t
                                                writer.RenderEndTag()
                                | Prefix "Project" rest -> 
                                                writer.RenderBeginTag HtmlTextWriterTag.Html
                                                writer.RenderBeginTag HtmlTextWriterTag.Head
                                                writer.RenderBeginTag HtmlTextWriterTag.Title
                                                writer.Write t
                                                writer.RenderEndTag()
                                                writer.RenderEndTag()
                                                writer.RenderBeginTag HtmlTextWriterTag.Body
                                                writer.RenderBeginTag HtmlTextWriterTag.H1
                                                writer.Write t
                                                writer.RenderEndTag()
                                | Prefix "Sat" rest -> 
                                                writer.Write t
                                | Prefix "Sun" rest -> writer.Write t
                                | _ -> 
                                        writer.RenderBeginTag HtmlTextWriterTag.P
                                        writer.Write t
                                        writer.RenderEndTag()
                                        )

    writer.RenderEndTag()
    writer.RenderEndTag()

    string stringWriter |> printfn "%s"
    string stringWriter

let HTMLWriter i (content: string) = 
    use streamWriter = new StreamWriter(BLOG_HTML_DESTINTION_PATH + "blog" + string (i+1) + ".html", false)
    streamWriter.Write content
    content
 
type Rss = XmlProvider<"https://a3qz.github.io/ethics/rss.xml">
let blog = Rss.GetSample()

let ItemBuilder (content: string) = 
    content




let ethicsbloghead = """<?xml version="1.0"?>
<rss version="2.0"
     xmlns:sy="http://purl.org/rss/1.0/modules/syndication/">
  <channel>
    <language>en</language>
    
    <sy:updatePeriod>hourly</sy:updatePeriod>
    <sy:updateFrequency>1</sy:updateFrequency>
    <title>a3qz Ethics Blog</title>
    <link>https://a3qz.github.io</link>
    <description>Blog for CSE 40175</description>
"""
let blogtail =   """</channel>
</rss>"""

let hocbloghead = """<?xml version="1.0"?>
<rss version="2.0"
     xmlns:sy="http://purl.org/rss/1.0/modules/syndication/">
  <channel>
    <language>en</language>
    
    <sy:updatePeriod>hourly</sy:updatePeriod>
    <sy:updateFrequency>1</sy:updateFrequency>
    <title>a3qz History of Computing Blog</title>
    <link>https://a3qz.github.io</link>
    <description>Blog for CSE 40850</description>
"""


let itemtext = """<item><title>$</title>
      <link>https://a3qz.github.io/ethics/blogs/blog#.html</link>
      <pubDate>&</pubDate></item>"""

let blogheads = Map.empty.Add("history", hocbloghead).Add("ethics",ethicsbloghead)

// Take a string of html with my stupid ~ as delimiter thing and split it, then write out the html 
let RSSMaker i (text: string) = 
    //let stringWriter = new StringWriter()
    use writer = new StreamWriter(SITE_RSS_DESTINATION_PATH, true)
    let i = i + 1 |> string
    text.Split [|'~'|]
    |> Seq.iter (fun (t:string) -> match (t:string) with 
                                | Prefix "Reading" rest -> 
                                                "<item><title>$</title>".Replace("$", t) |> writer.Write 
                                                
                                | Prefix "Project" rest -> 
                                                "<item><title>$</title>".Replace("$", t) |> writer.Write 
                                | Prefix "Sat" rest -> 
                                                let temp = "<link>https://a3qz.github.io/"+BLOG_CLASS+"/blogs/blog$.html</link>"
                                                temp.Replace("$", i) |> writer.Write
                                                "<pubDate>$</pubDate></item>".Replace("$", t) |> writer.Write 
                                | Prefix "Sun" rest -> 
                                                let temp = "<link>https://a3qz.github.io/"+BLOG_CLASS+"/blogs/blog$.html</link>"
                                                temp.Replace("$", i) |> writer.Write
                                                "<pubDate>$</pubDate></item>".Replace("$", t) |> writer.Write 
                                | _ -> 
                                        writer.Write ""
                                        )
    writer.Close()

let RSSWriter (content: string) = 
    use streamWriter = new StreamWriter(SITE_RSS_DESTINATION_PATH, true)
    streamWriter.Write content
    content

// Effectivley main; get all the files in the directory and call FileReader on them
Directory.GetFiles(BLOG_DOCUMENT_SOURCE_PATH, "*") 
    |> Array.map Path.GetFileName 
    |> Array.map FileReader
    |> Array.map HTMLMaker
    |> Array.mapi HTMLWriter
    |> Array.iter (printfn "%s")


let streamWritertemp = new StreamWriter(SITE_RSS_DESTINATION_PATH, false)
blogheads.[BLOG_CLASS] |>streamWritertemp.Write 
streamWritertemp.Close() |> ignore
Directory.GetFiles(BLOG_DOCUMENT_SOURCE_PATH, "*") 
    |> Array.map Path.GetFileName 
    |> Array.map FileReader
    |> Array.mapi RSSMaker
    |> ignore
let streamWritertemp2 = new StreamWriter(SITE_RSS_DESTINATION_PATH, true)
streamWritertemp2.Write blogtail
streamWritertemp2.Close() |> ignore
// Makes it so that I can read the output without the window closing
//System.Console.ReadLine() |> ignore
