using Terminal.Gui;

namespace CreateXslt.Views;

public class LandingPage
{
    public static Window GetLandingPage()
    {
        var Landingpage = new Window("Dansk metal Rapport generator")
        {
            X = 0,
            Y = 0,
            Width = Dim.Percent(100),
            Height = Dim.Fill() - 1,
            TextAlignment = TextAlignment.Centered,
            //TextFormatter = new TextFormatter() {Size = new Size(1, 10)}
        };
        
        //TODO: Add landing page info, how to
        
        return Landingpage;
    }
}