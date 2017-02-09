namespace NeXL.Bloomberg
open System
open System.Windows
open System.Windows.Controls
open System.Windows.Media
open System.IO
open System.Threading
open System.Collections.Generic
open System.Windows.Documents
open System.Windows.Navigation
open System.Windows.Media.Imaging
open System.Reflection
open NeXL.ManagedXll
open StateManager
open GlobalNames

[<XlInvisible>]
module AboutWindow =

    let private createRowDef (length : float) (unit : GridUnitType) =
        let rowDef = new RowDefinition()
        rowDef.Height <- new GridLength(length, unit)
        rowDef

    let private createColDef (length : float) (unit : GridUnitType) =
        let colDef = new ColumnDefinition()
        colDef.Width <- new GridLength(length, unit)
        colDef

    let private logo = EmbeddedResources.getBitmap "Logo.gif"

    let create() =
        let window = new Window()
        window.WindowStartupLocation <- WindowStartupLocation.CenterScreen
        window.Title <- "About"
        window.MinWidth <- 400.
        window.Background <- SystemColors.ControlBrush
        window.SizeToContent <- SizeToContent.WidthAndHeight
        window.ResizeMode <- ResizeMode.NoResize

        let grid = new Grid()
        [1..8] |> List.iter (fun i -> grid.RowDefinitions.Add(createRowDef 1.0 GridUnitType.Auto))
        [1..1] |> List.iter (fun i -> grid.ColumnDefinitions.Add(createColDef 1.0 GridUnitType.Auto))

        let image = new Image()
        image.HorizontalAlignment <- HorizontalAlignment.Left
        let bitmap = new BitmapImage()
        bitmap.BeginInit()
        use stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Logo.gif")
        bitmap.StreamSource <- stream
        bitmap.CacheOption <- BitmapCacheOption.OnLoad
        bitmap.EndInit()
        bitmap.Freeze()
        image.Source <- bitmap
        Grid.SetRow(image, 0)
        Grid.SetColumn(image, 0)

        let label1 = new TextBlock()
        label1.Margin <- new Thickness(5.)
        label1.Text <- "Excel Connector for Bloomberg API. Powered by NeXL Technology"
        label1.FontWeight <- FontWeights.Bold
        Grid.SetRow(label1, 1)
        Grid.SetColumn(label1, 0)

        let label2 = new TextBlock()
        label2.Margin <- new Thickness(5.)
        label2.Text <- sprintf "Version %d.%d" (Assembly.GetExecutingAssembly().GetName().Version.Major) (Assembly.GetExecutingAssembly().GetName().Version.Minor) 
        Grid.SetRow(label2, 2)
        Grid.SetColumn(label2, 0)

        let label3 = new TextBlock()
        label3.Margin <- new Thickness(5.)
        label3.Text <- "©StatFactory Ltd UK 2017"
        Grid.SetRow(label3, 3)
        Grid.SetColumn(label3, 0)

        let label4 = new Label()
        let run = new Run("Source Code on GitHub")
        let link = new Hyperlink(run)
        label4.Content <- link
        link.RequestNavigate.Add(fun args ->         
                                    try
                                        System.Diagnostics.Process.Start("https://github.com/Statfactory/NeXL.Bloomberg") |> ignore
                                    with
                                        | _ -> ())
        link.NavigateUri <- new Uri("https://github.com/Statfactory/NeXL.Bloomberg")
        Grid.SetRow(label4, 4)
        Grid.SetColumn(label4, 2)

        let label5 = new TextBlock()
        label5.Margin <- new Thickness(5.)
        label5.Text <- "License:"
        label5.FontWeight <- FontWeights.Bold
        Grid.SetRow(label5, 5)
        Grid.SetColumn(label5, 0)

        let label6 = new Label()
        let runLic = new Run("Creative Commons 4.0")
        let linkLic = new Hyperlink(runLic)
        label6.Content <- linkLic
        linkLic.RequestNavigate.Add(fun args ->         
                                        try
                                            System.Diagnostics.Process.Start("https://creativecommons.org/licenses/by-nd/4.0/legalcode") |> ignore
                                        with
                                            | _ -> ())
        linkLic.NavigateUri <- new Uri("https://creativecommons.org/licenses/by-nd/4.0/legalcode")
        Grid.SetRow(label6, 6)
        Grid.SetColumn(label6, 0)

        let label7 = new TextBlock()
        label7.Margin <- new Thickness(5.)
        label7.Text <- "Contact: nexl@statfactory.co.uk"
        Grid.SetRow(label7, 7)
        Grid.SetColumn(label7, 0)

        grid.Children.Add(image) |> ignore
        grid.Children.Add(label1) |> ignore
        grid.Children.Add(label2) |> ignore
        grid.Children.Add(label3) |> ignore
        grid.Children.Add(label4) |> ignore
        grid.Children.Add(label5) |> ignore
        grid.Children.Add(label6) |> ignore
        grid.Children.Add(label7) |> ignore

        window.Content <- grid
        window
