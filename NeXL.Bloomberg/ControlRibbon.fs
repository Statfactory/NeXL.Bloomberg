namespace NeXL.Bloomberg
open System
open NeXL.ManagedXll
open System.Runtime.InteropServices
open System.Windows.Forms
open System.Drawing
open System.Reflection
open NeXL.XlInterop

[<ComVisible(true)>]
[<XlInvisibleAttribute>]
type ControlRibbon() =
    inherit XlRibbon()

    let mutable ribbon = null

    override this.GetCustomUI(ribbonID : string) =
        @"
        <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoad' loadImage='LoadImage'>
            <ribbon>
            <tabs>
                <tab id='TabId' label='NeXL.Bloomberg'>
                <group id='GroupId' label='Control'>
                    <button id='ButtonInstallId' image='ExcelAddIn' label='Activate As Excel AddIn' size='large' getEnabled='getInstallEnabled' onAction='OnActivateButtonClick' />
                    <button id='ButtonDeleteId' imageMso='Delete' label='Remove Excel AddIn' size='large' getEnabled='getDeleteEnabled' onAction='OnDeleteButtonClick' />
                    <button id='ButtonAboutId' image='About' label='About' size='large' onAction='OnAboutButtonClick' />
                </group>
                </tab>
            </tabs>
            </ribbon>
        </customUI>"

    override this.OnLoad(ribbonUI : IRibbonUI) =
        ribbon <- ribbonUI

    member this.getInstallEnabled(control : IRibbonControl ) =
        ExcelCustomization.CustomizationIsInstalled() |> not

    member this.getDeleteEnabled(control : IRibbonControl ) =
        ExcelCustomization.CustomizationIsInstalled()

    member this.OnActivateButtonClick(control : IRibbonControl) =
        ExcelCustomization.DeleteCustomization()
        ExcelCustomization.Install()
        if ribbon <> null then ribbon.Invalidate()

    member this.OnDeleteButtonClick(control : IRibbonControl) =
        ExcelCustomization.DeleteCustomization()
        if ribbon <> null then ribbon.Invalidate()

    member this.OnAboutButtonClick(control : IRibbonControl) =
        let aboutWindow = AboutWindow.create()
        aboutWindow.ShowDialog() |> ignore

    override this.GetBitmap(imageName) =
        use stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(imageName + ".png")
        Bitmap.FromStream(stream):?>Bitmap
        


