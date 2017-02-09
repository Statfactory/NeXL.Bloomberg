namespace NeXL.Bloomberg
open System
open NeXL.ManagedXll
open System.Runtime.InteropServices
open System.Windows.Forms
open System.Reflection
open NeXL.XlInterop

module CustomizationProviders =

    [<XlWorkbookTypesProvider>]
    let getWorkbookTypes(workbook : obj) = [|typeof<ControlRibbon>|]

