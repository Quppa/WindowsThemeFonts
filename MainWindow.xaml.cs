namespace WindowsThemeFonts
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;
    using System.Windows.Documents;
    using System.Windows.Input;
    using System.Windows.Interop;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using System.Windows.Navigation;
    using System.Windows.Shapes;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("uxtheme.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr OpenThemeData(IntPtr hWnd, String classList);

        [DllImport("uxtheme.dll", ExactSpelling = true)]
        public extern static Int32 CloseThemeData(IntPtr hTheme);

        [DllImport("uxtheme", ExactSpelling = true, CharSet = CharSet.Unicode)]
        public extern static Int32 GetThemeFont(IntPtr hTheme, IntPtr hdc, int iPartId, int iStateId, int iPropId, out LOGFONT pFont);

        [DllImport("uxtheme", ExactSpelling = true)]
        public extern static Int32 GetThemeColor(IntPtr hTheme, int iPartId, int iStateId, int iPropId, out COLORREF pColor);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct LOGFONT
        {
            public const int LF_FACESIZE = 32;
            public int lfHeight;
            public int lfWidth;
            public int lfEscapement;
            public int lfOrientation;
            public int lfWeight;
            public byte lfItalic;
            public byte lfUnderline;
            public byte lfStrikeOut;
            public byte lfCharSet;
            public byte lfOutPrecision;
            public byte lfClipPrecision;
            public byte lfQuality;
            public byte lfPitchAndFamily;
            [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = LF_FACESIZE)]
            public string lfFaceName;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct COLORREF
        {
            public byte R;
            public byte G;
            public byte B;
        }

        public enum TEXTSTYLEPARTS
        {
            TEXT_MAININSTRUCTION = 1,
            TEXT_INSTRUCTION = 2,
            TEXT_BODYTITLE = 3,
            TEXT_BODYTEXT = 4,
            TEXT_SECONDARYTEXT = 5,
            TEXT_HYPERLINKTEXT = 6,
            TEXT_EXPANDED = 7,
            TEXT_LABEL = 8,
            TEXT_CONTROLLABEL = 9,
        };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Update();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Update();
        }

        private void Update()
        {
            IntPtr hTheme = IntPtr.Zero;

            try
            {
                FontFamily fontfamily = null;
                double fontsizediu = 0, fontsizepx = 0;
                FontWeight fontweight = FontWeights.Normal;
                Brush foreground = null;

                // get a handle to theme data for textstyle
                hTheme = OpenThemeData(IntPtr.Zero, "TEXTSTYLE");

                bool fallbackfont = false, fallbackcolour = false;

                // if hTheme is null, 'TEXTSTYLE' could not be found - this probably means theming is turned off
                if (hTheme == IntPtr.Zero)
                {
                    fallbackfont = true;
                    fallbackcolour = true;
                }
                else
                {
                    int S_OK = 0;

                    // from vssym32.h
                    int TMT_TEXTCOLOR = 3803;

                    COLORREF pColor;
                    if (GetThemeColor(hTheme, (int)TEXTSTYLEPARTS.TEXT_MAININSTRUCTION, 0, TMT_TEXTCOLOR, out pColor) == S_OK)
                    {
                        foreground = new SolidColorBrush(Color.FromRgb(pColor.R, pColor.G, pColor.B));
                    }
                    else
                    {
                        // fall back to the window text colour
                        fallbackcolour = true;
                    }

                    // from vssym32.h
                    int TMT_FONT = 210;

                    LOGFONT lFont;
                    if (GetThemeFont(hTheme, IntPtr.Zero, (int)TEXTSTYLEPARTS.TEXT_MAININSTRUCTION, 0, TMT_FONT, out lFont) == S_OK)
                    {
                        fontfamily = new FontFamily(lFont.lfFaceName);
                        fontweight = FontWeight.FromOpenTypeWeight(lFont.lfWeight);

                        //// important: lfHeight contains the height in pixels
                        //// e.g. for main instruction text in Aero at 96dpi is 16px high, and at 120dpi is 20px high
                        //// WPF uses device independent units, however - 16 DIU = 16px @ 96dpi, 20px @ 120dpi, etc.

                        FontSizeConverter converter = new FontSizeConverter();
                        fontsizepx = Math.Abs(lFont.lfHeight);
                    }
                    else
                    {
                        // fall back to caption font
                        fallbackfont = true;
                    }
                }

                if (fallbackfont)
                {
                    // if there is no theming, use SystemFonts.CaptionFont*
                    // also fall back to this if GetThemeFont fails
                    fontfamily = SystemFonts.CaptionFontFamily;
                    fontsizediu = SystemFonts.CaptionFontSize;
                    fontweight = SystemFonts.CaptionFontWeight;
                }

                if (fallbackcolour)
                {
                    // if there is no theming, use SystemColors.WindowTextBrush
                    // also fall back to this if GetThemeColor fails
                    foreground = SystemColors.WindowTextBrush;
                }

                #region DPI Calculations

                // get the handle of the window
                HwndSource windowhandlesource = PresentationSource.FromVisual(this) as HwndSource;

                // work out the current screen's DPI
                Matrix screenmatrix = windowhandlesource.CompositionTarget.TransformToDevice;
                //double dpiX = screenmatrix.M11; 
                double dpiY = screenmatrix.M22; // 1.0 = 96dpi, 1.25 = 120dpi, etc.

                if (fontsizepx == 0)
                {
                    // convert from DIUs to pixels
                    fontsizepx = fontsizediu * dpiY;
                }
                else if (fontsizediu == 0)
                {
                    // convert from pixels to DIUs
                    fontsizediu = fontsizepx / dpiY;
                }

                #endregion

                // set text block style
                MainInstructionTextBlock.FontFamily = fontfamily;
                MainInstructionTextBlock.FontSize = fontsizediu;
                MainInstructionTextBlock.FontWeight = fontweight;
                MainInstructionTextBlock.Foreground = foreground;

                // display info
                MainInstructionFontFamilyTextBlock.Text = fontfamily.ToString();
                MainInstructionFontSizeTextBlock.Text = fontsizediu.ToString() + " DIU (" + fontsizepx + " px)";
                MainInstructionFontWeightTextBlock.Text = fontweight.ToString();
                MainInstructionForegroundTextBlock.Text = foreground.ToString();

                MainInstructionFontFallbackTextBlock.Text = fallbackfont.ToString();
                MainInstructionColourFallbackTextBlock.Text = fallbackcolour.ToString();
            }
            finally
            {
                // clean up
                if (hTheme != IntPtr.Zero)
                    CloseThemeData(hTheme);
            }
        }
    }
}
