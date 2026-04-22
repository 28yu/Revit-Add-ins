using System.Collections.Generic;

namespace Tools28.Commands.FormworkCalculator.Models
{
    public enum FaceType
    {
        FormworkRequired,
        DeductedTop,
        DeductedBottom,
        DeductedContact,
        DeductedBelowGL,
        Inclined,
        Error,
    }

    public enum CategoryGroup
    {
        Column,
        Beam,
        Wall,
        Slab,
        Foundation,
        Stairs,
        Other,
    }

    public class ElementResult
    {
        public int ElementId { get; set; }
        public string ElementName { get; set; } = string.Empty;
        public CategoryGroup Category { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string Zone { get; set; } = string.Empty;
        public string FormworkType { get; set; } = string.Empty;
        public double FormworkArea { get; set; }
        public double DeductedTopArea { get; set; }
        public double DeductedBottomArea { get; set; }
        public double DeductedContactArea { get; set; }
        public double InclinedArea { get; set; }
        public double OpeningAreaDeducted { get; set; }
        public double OpeningEdgeAreaAdded { get; set; }
    }

    public class FaceAnalysisResult
    {
        public int SourceElementId { get; set; }
        public FaceType FaceType { get; set; }
        public double Area { get; set; }
        public List<Autodesk.Revit.DB.CurveLoop> BoundaryLoops { get; set; }
            = new List<Autodesk.Revit.DB.CurveLoop>();
        public Autodesk.Revit.DB.XYZ Normal { get; set; }
        public string GroupKey { get; set; } = string.Empty;
    }

    public class CategoryResult
    {
        public CategoryGroup Category { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public double FormworkArea { get; set; }
        public double DeductedArea { get; set; }
        public int ElementCount { get; set; }
    }

    public class ZoneResult
    {
        public string Zone { get; set; } = string.Empty;
        public double FormworkArea { get; set; }
        public int ElementCount { get; set; }
    }

    public class FormworkTypeResult
    {
        public string FormworkType { get; set; } = string.Empty;
        public double FormworkArea { get; set; }
        public int ElementCount { get; set; }
    }

    public class ErrorLogEntry
    {
        public int ElementId { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string ElementName { get; set; } = string.Empty;
        public string ErrorKind { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }

    public class FormworkResult
    {
        public double TotalFormworkArea { get; set; }
        public double TotalDeductedArea { get; set; }
        public double InclinedFaceArea { get; set; }
        public int ProcessedElementCount { get; set; }

        public List<ElementResult> ElementResults { get; set; } = new List<ElementResult>();
        public List<CategoryResult> CategoryResults { get; set; } = new List<CategoryResult>();
        public List<ZoneResult> ZoneResults { get; set; } = new List<ZoneResult>();
        public List<FormworkTypeResult> TypeResults { get; set; } = new List<FormworkTypeResult>();
        public List<FaceAnalysisResult> FaceResults { get; set; } = new List<FaceAnalysisResult>();
        public List<ErrorLogEntry> Errors { get; set; } = new List<ErrorLogEntry>();
    }
}
