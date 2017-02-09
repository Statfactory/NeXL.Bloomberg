namespace NeXL.Bloomberg
open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Collections.Generic
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open Bloomberglp.Blpapi

type Agent<'T> = MailboxProcessor<'T>

type ObservableSubscription() as this =
    let mutable _observer : IObserver<Message> option = None

    let mutable correlationId : int64 option = None

    let mutable session : Session option = None

    interface IObservable<Message> with  
        member __.Subscribe(observer : IObserver<Message>) : IDisposable = 
            match _observer with 
                | None -> _observer <- Some observer
                | _ -> ()
            this:>IDisposable

    interface IDisposable with
        member __.Dispose() =
            match correlationId, session with
                | Some(correlationId), Some(session) ->
                    let sl = new List<Subscription>()
                    sl.Add(new Subscription("", new CorrelationID(correlationId)))
                    session.Unsubscribe(sl)
                    _observer <- None
                | _ -> ()

    member this.CorrelationId   
        with set(v) = correlationId <- Some v

    member this.Session   
        with set(v) = session <- Some v

    member this.SendEvent(el : Message) =
        match _observer with
            | Some(observer) -> observer.OnNext(el)
            | None -> ()

    member this.SendError(exn) =
         match _observer with
            | Some(observer) ->
                observer.OnError(exn)
                _observer <- None
            | None -> ()       

type SessionMsg =
    | OpenSession of serverHost : string * serverPort : int
    | SessionEvent of event : Event * session : Session
    | RefDataRequest of securities : string[] * fields : string[] * reply : AsyncReplyChannel<Element list>
    | SubscriptionStart of topic : string * fields : string[] * obsSubscription : ObservableSubscription

type SessionState =
    {
        Session : Session option
        RefDataService : Service option
        LastCorrId : int64
        RequestResponses : Dictionary<int64, Element list>
        Subscriptions : Dictionary<int64, ObservableSubscription>
        Replies : Dictionary<int64, AsyncReplyChannel<Element list>>
    }

type StatusEvent =
    {
        Status : string
    }

[<XlInvisible>]
type BlpSession(serverHost : string, serverPort : int) =
    let refDataSvcCorrId = 1L

    let statusEvent = new Event<StatusEvent>()
    let processRefDataResponse (messages : seq<Message>) (responses : Dictionary<int64, Element list>) =
        messages |> Seq.filter (fun msg -> msg.MessageType.Equals("ReferenceDataResponse") && msg.AsElement.HasElement("securityData"))
                 |> Seq.collect (fun msg ->
                                    let secDataArr = msg.AsElement.GetElement("securityData")
                                    seq{for i in 0..secDataArr.NumValues - 1 do yield msg.CorrelationID.Value, secDataArr.GetValueAsElement(i)}
                                    |> Seq.filter (fun (corrId, el) -> el.HasElement("fieldData"))
                                )
                 |> Seq.groupBy fst
                 |> Seq.iter (fun (corrId, elems) -> 
                                let elems = elems |> Seq.map snd |> Seq.toList
                                if responses.ContainsKey(corrId) then
                                    let currElems = responses.[corrId]
                                    responses.[corrId] <- currElems @ elems
                                else
                                    responses.Add(corrId, elems)
                            )


    let agent =
        Agent.Start(fun inbox ->
                        let rec msgLoop(state : SessionState) =
                            async
                                {
                                 let! msg = inbox.Receive()
                                 match msg with    
                                    | OpenSession(serverHost, serverPort) ->
                                        let sessionOptions = new SessionOptions(ServerHost = serverHost, ServerPort = serverPort)
                                        let session = new Session(sessionOptions, new EventHandler(fun e s -> inbox.Post(SessionEvent(e, s))))
                                        session.StartAsync() |> ignore
                                        statusEvent.Trigger({Status = "Opening session"})
                                        return! msgLoop(state) 

                                    | SessionEvent(ev, s) ->
                                        match ev.Type with
                                            | Event.EventType.SESSION_STATUS ->
                                                let msgs = ev.GetMessages()
                                                if msgs |> Seq.exists (fun msg -> msg.MessageType.Equals("SessionStarted")) then 
                                                    statusEvent.Trigger({Status = "Session started"})
                                                    s.OpenServiceAsync("//blp/refdata", new CorrelationID(refDataSvcCorrId))
                                                    return! msgLoop({state with Session = Some s}) 
                                                elif msgs |> Seq.exists (fun msg -> msg.MessageType.Equals("SessionTerminated")) then 
                                                    statusEvent.Trigger({Status = "Session terminated"})
                                                    return! msgLoop({state with Session = None}) 
                                                else
                                                    return! msgLoop({state with Session = Some s}) 

                                            | Event.EventType.SERVICE_STATUS ->
                                                let msgs = ev.GetMessages()
                                                msgs |> Seq.iter (fun m -> statusEvent.Trigger({Status = m.MessageType.ToString()}))
                                                match msgs |> Seq.tryFind (fun msg -> msg.MessageType.Equals("ServiceOpened")) with
                                                    | Some(msg) when msg.CorrelationID.Value = refDataSvcCorrId -> 
                                                        let service = s.GetService("//blp/refdata")
                                                        statusEvent.Trigger({Status = "Reference data service opened"})
                                                        return! msgLoop({state with RefDataService = Some service}) 
                                                    | _ -> return! msgLoop(state) 

                                            | Event.EventType.PARTIAL_RESPONSE ->
                                                let msgs = ev.GetMessages()
                                                processRefDataResponse msgs state.RequestResponses
                                                return! msgLoop(state) 

                                            | Event.EventType.RESPONSE ->
                                                let msgs = ev.GetMessages()
                                                processRefDataResponse msgs state.RequestResponses
                                                let corrIds = msgs |> Seq.map (fun m -> m.CorrelationID.Value) |> Seq.distinct 
                                                corrIds |> Seq.iter (fun corrId -> 
                                                                         if state.Replies.ContainsKey(corrId) && state.RequestResponses.ContainsKey(corrId) then
                                                                             let reply = state.Replies.[corrId]
                                                                             let elems = state.RequestResponses.[corrId]
                                                                             reply.Reply(elems)
                                                                             state.Replies.Remove(corrId) |> ignore
                                                                             state.RequestResponses.Remove(corrId) |> ignore
                                                                    )
                                                return! msgLoop(state) 

                                            | Event.EventType.SUBSCRIPTION_STATUS ->
                                                let msgs = ev.GetMessages()
                                                msgs |> Seq.filter (fun msg -> msg.MessageType.Equals("SubscriptionFailure") && msg.HasElement("SubscriptionFailure"))
                                                     |> Seq.iter (fun msg -> 
                                                                      let corrId = msg.CorrelationID.Value
                                                                      if state.Subscriptions.ContainsKey(corrId) then  
                                                                          let s = state.Subscriptions.[corrId]
                                                                          let errMsg = msg.AsElement.GetElement("SubscriptionFailure").GetElement("reason").GetElementAsString("description")
                                                                          s.SendError(new ArgumentException(errMsg))
                                                                          statusEvent.Trigger({Status = sprintf "Subscription failed; %s" errMsg})
                                                                          state.Subscriptions.Remove(corrId) |> ignore
                                                                 )
                                                return! msgLoop(state)

                                            | Event.EventType.SUBSCRIPTION_DATA ->
                                                let msgs = ev.GetMessages()
                                                msgs |> Seq.filter (fun msg -> msg.MessageType.Equals("MarketDataEvents"))
                                                     |> Seq.map (fun msg -> msg.CorrelationID.Value, msg)
                                                     |> Seq.iter (fun (corrId, elem) ->
                                                                      if state.Subscriptions.ContainsKey(corrId) then
                                                                          state.Subscriptions.[corrId].SendEvent(elem)
                                                                  )
                                                return! msgLoop(state)

                                            | _ -> return! msgLoop(state)

                                    | RefDataRequest(secs, fields, reply) -> 
                                        match state.Session, state.RefDataService with   
                                            | Some(session), Some(svc) ->
                                                try
                                                    let corrId = state.LastCorrId + 1L
                                                    let request = svc.CreateRequest("ReferenceDataRequest")

                                                    secs |> Array.iter (fun s -> request.Append("securities", s))
                                                    fields |> Array.iter (fun f -> request.Append("fields", f))

                                                    session.SendRequest(request, new CorrelationID(corrId)) |> ignore

                                                    state.Replies.Add(corrId, reply)

                                                    return! msgLoop({state with LastCorrId = corrId}) 
                                                with ex ->
                                                    statusEvent.Trigger({Status = ex.StackTrace})
                                                    return! msgLoop(state)


                                            | _ ->
                                                do! Async.Sleep(200)
                                                inbox.Post(msg)
                                                return! msgLoop(state) 

                                    | SubscriptionStart(topic, fields, obsSubscription) ->
                                        statusEvent.Trigger({Status = "Subscription starting"})
                                        match state.Session with   
                                            | Some(session) ->
                                                let corrId = state.LastCorrId + 1L
                                                obsSubscription.CorrelationId <- corrId
                                                obsSubscription.Session <- session

                                                let sl = new List<Subscription>()
                                                sl.Add(new Subscription(topic, fields, new CorrelationID(corrId)))
                                                state.Subscriptions.Add(corrId, obsSubscription)
                                                session.Subscribe(sl)
                                                return! msgLoop({state with LastCorrId = corrId}) 

                                            | _ ->
                                                do! Async.Sleep(200)
                                                inbox.Post(msg)
                                                return! msgLoop(state) 

                                }
                        msgLoop({Session = None; RefDataService = None; LastCorrId = 2L
                                 RequestResponses = new Dictionary<_,_>(); Subscriptions = new Dictionary<_,_>()
                                 Replies = new Dictionary<_,_>()})
                   )

    do 
        agent.Post(OpenSession(serverHost, serverPort))
        

    new() = new BlpSession("localhost", 8194)

    member this.SendRefDataRequest(secs : string[], fields : string[], timeout) =
        agent.PostAndTryAsyncReply((fun reply -> RefDataRequest(secs, fields, reply)), timeout)

    member this.StartSubscription(topic : string, fields : string[], obsSubscription) =
        agent.Post(SubscriptionStart(topic, fields, obsSubscription))

    member this.StatusEvent = statusEvent.Publish
