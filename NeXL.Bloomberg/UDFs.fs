namespace NeXL.Bloomberg
open NeXL.ManagedXll
open NeXL.XlInterop
open NeXL.ManagedXll.ExplicitConversion
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open Bloomberglp.Blpapi

[<XlQualifiedName(true)>]
module Blp =
    let private defaultSession = lazy (new BlpSession())

    let private elementToXlValue (el : Element) (field : string) =
        match seq  
                {
                    yield try Some(XlNumeric(float(el.GetElementAsInt32(field)))) with _ -> None
                    yield try Some(XlNumeric(el.GetElementAsFloat64(field))) with _ -> None
                    yield try Some(XlNumeric(el.GetElementAsDate(field).ToSystemDateTime().ToOADate())) with _ -> None
                    yield try Some(XlNumeric(el.GetElementAsTime(field).ToSystemDateTime().ToOADate())) with _ -> None
                    yield try Some(XlNumeric(el.GetElementAsDatetime(field).ToSystemDateTime().ToOADate())) with _ -> None
                    yield try Some(XlString(el.GetElementAsString(field))) with _ -> None
                } |> Seq.tryFind (Option.isSome) with
            | Some(Some(v)) -> v
            | _ -> XlNil 

    let private messageToXlValue (field : string) (msg : Message) =
        match seq  
                {
                    yield try Some(XlNumeric(float(msg.GetElementAsInt32(field)))) with _ -> None
                    yield try Some(XlNumeric(msg.GetElementAsFloat64(field))) with _ -> None
                    yield try Some(XlNumeric(msg.GetElementAsDate(field).ToSystemDateTime().ToOADate())) with _ -> None
                    yield try Some(XlNumeric(msg.GetElementAsTime(field).ToSystemDateTime().ToOADate())) with _ -> None
                    yield try Some(XlNumeric(msg.GetElementAsDatetime(field).ToSystemDateTime().ToOADate())) with _ -> None
                    yield try Some(XlString(msg.GetElementAsString(field))) with _ -> None
                } |> Seq.tryFind (Option.isSome) with
            | Some(v) -> v
            | _ -> None

    let private refDataToXlTable (fields : string[]) (elems : Element list) : XlTable =
        match elems with
            | [] -> XlTable.Empty
            | h :: t ->
                let cols = Array.init (fields.Length + 1) (fun i -> if i = 0 then {Name = "Security"; IsDateTime = false} else {Name = fields.[i - 1]; IsDateTime = false})
                let data = Array2D.init (elems.Length) (fields.Length + 1) 
                                        (fun i j -> 
                                             if j = 0 then 
                                                 elementToXlValue elems.[i] "security"  
                                             else
                                                 elementToXlValue (elems.[i].GetElement("fieldData")) fields.[j - 1]
                                        )
                new XlTable(cols, data, "", "", false, false, true)

    let private marketDataToXlTable (fields : string[]) (msgs : Message list) : XlTable =
        let cols = Array.init (fields.Length) (fun i -> {Name = fields.[i]; IsDateTime = false})
        let data = Array2D.init 1 cols.Length (fun i j ->
                                                   let field = fields.[j]
                                                   match msgs |> Seq.map (messageToXlValue field) |> Seq.tryFind Option.isSome with
                                                       | Some(Some(v)) -> v
                                                       | _ -> XlNil
                                               )
        new XlTable(cols, data, "", "", false, false, true)

    let createSession (serverHost : string, serverPort : int) = new BlpSession(serverHost, serverPort)

    let getRefData (securities : string[], fields : string[], session : BlpSession option, timeout : int option) =
        let session = defaultArg session defaultSession.Value
        let timeout = defaultArg timeout 5000
        async  
            {
             let! respElems = session.SendRefDataRequest(securities, fields, timeout)
             match respElems with
                 | Some(elems) -> return refDataToXlTable fields elems
                 | None -> return new XlTable(XlString("Timeout or bad request"))
            }

    let getMarketData (topic : string, fields : string[], session : BlpSession option) =
        let fields = fields |> Array.distinct
        let session = defaultArg session defaultSession.Value
        let subscription = new ObservableSubscription()
        session.StartSubscription(topic, fields, subscription) 
        subscription |> Observable.map (marketDataToXlTable fields)

    let getSessionEvents (session : BlpSession option) =
        let session = defaultArg session defaultSession.Value
        session.StatusEvent |> Event.scan (fun s e -> e :: s) []
                            |> Event.map (List.rev >> List.toArray)
                            |> Event.map (fun e -> XlTable.Create(e, "", "", false, false, true))


    let getErrors newOnTop =
        UdfErrorHandler.OnError |> Event.scan (fun s e -> e :: s) []
                                |> Event.map (fun errs ->
                                                  let errs = if newOnTop then errs |> List.toArray else errs |> List.rev |> List.toArray
                                                  XlTable.Create(errs, "", "", false, false, true)
                                             )





