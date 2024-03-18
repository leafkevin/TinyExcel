namespace TinyExcel;

public class SpreadsheetDocument
{
    public ApplicationType ApplicationType => ApplicationType.Excel;
    public string Name { get; set; }
    public string RelationshipType => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";   
}
