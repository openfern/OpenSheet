using XL = NPOI.SS.UserModel;

namespace OpenSheet.Printing
{
    public enum PrintOrientation { Default = 1, Portrait = 2, Landscape = 3 }

    /// <summary>
    /// Paper sizes. The numeric values are the library's own and are NOT the codes stored in the file (they skip
    /// "Letter Small", so <c>A4</c> is 8 here but Excel's code for A4 is 9). The names are mapped to the real codes
    /// when reading and writing; see <see cref="PaperCodes"/>.
    /// </summary>
    public enum PaperSize
    {
        PrinterDefault = 0, Letter = 1, Tabloid = 2, Ledger = 3, Legal = 4, Statement = 5, Executive = 6, A3 = 7, A4 = 8,
        A4Small = 9, A5 = 10, B4 = 11, B5 = 12, Folio = 13, Quarto = 14, TenByFourteen = 15, ElevenBySeven = 16, Note = 17,
        Envelope9 = 18, Envelope10 = 19, Envelope11 = 20, Envelope12 = 21, Envelope14 = 22, CSizeSheet = 23, DSizeSheet = 24,
        ESizeSheet = 25, EnvelopeDL = 26, EnvelopeC5 = 27, EnvelopeC3 = 28, EnvelopeC4 = 29, EnvelopeC6 = 30,
        EnvelopeMonarch = 31, A4EExtra = 53, A4Transverse = 55, A4Plus = 60, LetterRotated = 75, A4Rotated = 77,
    }

    /// <summary>Maps <see cref="PaperSize"/> names to the paper codes Excel stores (the Windows DMPAPER numbers).</summary>
    static class PaperCodes
    {
        static readonly (PaperSize size, short code)[] Table =
        {
            (PaperSize.PrinterDefault, 0), (PaperSize.Letter, 1), (PaperSize.Tabloid, 3), (PaperSize.Ledger, 4), (PaperSize.Legal, 5),
            (PaperSize.Statement, 6), (PaperSize.Executive, 7), (PaperSize.A3, 8), (PaperSize.A4, 9), (PaperSize.A4Small, 10),
            (PaperSize.A5, 11), (PaperSize.B4, 12), (PaperSize.B5, 13), (PaperSize.Folio, 14), (PaperSize.Quarto, 15),
            (PaperSize.TenByFourteen, 16), (PaperSize.ElevenBySeven, 17), (PaperSize.Note, 18), (PaperSize.Envelope9, 19),
            (PaperSize.Envelope10, 20), (PaperSize.Envelope11, 21), (PaperSize.Envelope12, 22), (PaperSize.Envelope14, 23),
            (PaperSize.CSizeSheet, 24), (PaperSize.DSizeSheet, 25), (PaperSize.ESizeSheet, 26), (PaperSize.EnvelopeDL, 27),
            (PaperSize.EnvelopeC5, 28), (PaperSize.EnvelopeC3, 29), (PaperSize.EnvelopeC4, 30), (PaperSize.EnvelopeC6, 31),
            (PaperSize.EnvelopeMonarch, 37), (PaperSize.A4EExtra, 53), (PaperSize.A4Transverse, 55), (PaperSize.A4Plus, 60),
            (PaperSize.LetterRotated, 75), (PaperSize.A4Rotated, 77),
        };

        public static short ToCode(PaperSize size)
        {
            foreach (var (s, c) in Table) if (s == size) return c;
            throw new System.ArgumentOutOfRangeException(nameof(size), size, "Unknown paper size.");
        }

        public static PaperSize FromCode(short code)
        {
            foreach (var (s, c) in Table) if (c == code) return s;
            return (PaperSize)code; // a size this enum has no name for: keep the number
        }
    }

    public interface IPrintSetup
    {
        PrintOrientation PrintOrientation { get; set; }
        short Scale { get; set; }
        short PageStart { get; set; }
        PaperSize PaperSize { get; set; }
        short FitWidth { get; set; }
        short FitHeight { get; set; }
        bool LeftToRight { get; set; }
        bool NoColor { get; set; }
        bool Draft { get; set; }
        bool Notes { get; set; }
        bool UsePage { get; set; }
        short HResolution { get; set; }
        short VResolution { get; set; }
        double HeaderMargin { get; set; }
        double FooterMargin { get; set; }
        short Copies { get; set; }
    }

    internal sealed class PrintSetupAdapter : IPrintSetup
    {
        readonly XL.IPrintSetup _p;
        public PrintSetupAdapter(XL.IPrintSetup p) { _p = p; }

        public PrintOrientation PrintOrientation
        {
            get => _p.NoOrientation ? PrintOrientation.Default : _p.Landscape ? PrintOrientation.Landscape : PrintOrientation.Portrait;
            set
            {
                _p.NoOrientation = value == PrintOrientation.Default;
                _p.Landscape = value == PrintOrientation.Landscape;
            }
        }

        public short Scale { get => _p.Scale; set => _p.Scale = value; }
        public short PageStart { get => _p.PageStart; set => _p.PageStart = value; }
        public PaperSize PaperSize { get => PaperCodes.FromCode(_p.PaperSize); set => _p.PaperSize = PaperCodes.ToCode(value); }
        public short FitWidth { get => _p.FitWidth; set => _p.FitWidth = value; }
        public short FitHeight { get => _p.FitHeight; set => _p.FitHeight = value; }
        public bool LeftToRight { get => _p.LeftToRight; set => _p.LeftToRight = value; }
        public bool NoColor { get => _p.NoColor; set => _p.NoColor = value; }
        public bool Draft { get => _p.Draft; set => _p.Draft = value; }
        public bool Notes { get => _p.Notes; set => _p.Notes = value; }
        public bool UsePage { get => _p.UsePage; set => _p.UsePage = value; }
        public short HResolution { get => _p.HResolution; set => _p.HResolution = value; }
        public short VResolution { get => _p.VResolution; set => _p.VResolution = value; }
        public double HeaderMargin { get => _p.HeaderMargin; set => _p.HeaderMargin = value; }
        public double FooterMargin { get => _p.FooterMargin; set => _p.FooterMargin = value; }
        public short Copies { get => _p.Copies; set => _p.Copies = value; }
    }
}
