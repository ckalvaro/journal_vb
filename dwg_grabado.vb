' NX 2212
' Journal created by juanpablo on Thu Dec 28 16:06:51 2023 Hora estándar de Argentina

'
Imports System
Imports NXOpen

Module NXJournal
Sub Main (ByVal args() As String) 

Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
Dim workPart As NXOpen.Part = theSession.Parts.Work

Dim displayPart As NXOpen.Part = theSession.Parts.Display
Dim currentFile As String
Dim strRevision As String
Dim originalFile As String

' ----------------------------------------------
'   Menú: Archivo->Exportar->AutoCAD DXF/DWG...
' ----------------------------------------------


Dim dxfdwgCreator1 As NXOpen.DxfdwgCreator = Nothing
dxfdwgCreator1 = theSession.DexManager.CreateDxfdwgCreator()

dxfdwgCreator1.ExportData = NXOpen.DxfdwgCreator.ExportDataOption.Drawing

dxfdwgCreator1.AutoCADRevision = NXOpen.DxfdwgCreator.AutoCADRevisionOptions.R2004

dxfdwgCreator1.ViewEditMode = True

dxfdwgCreator1.FlattenAssembly = True

dxfdwgCreator1.ExportScaleValue = 1.0

dxfdwgCreator1.SettingsFile = "C:\splm\NX2212\dxfdwg\dxfdwg.def"

dxfdwgCreator1.OutputFileType = NXOpen.DxfdwgCreator.OutputFileTypeOption.Dwg

dxfdwgCreator1.ObjectTypes.Curves = True

dxfdwgCreator1.ObjectTypes.Annotations = True

dxfdwgCreator1.ObjectTypes.Structures = True

'dxfdwgCreator1.InputFile = "@DB/104584/A"
currentFile = workPart.GetStringAttribute("DB_PART_NO")
strRevision = workPart.GetStringAttribute("DB_PART_REV")
originalFile = "@DB/" & currentFile & "/" & strRevision
dxfdwgCreator1.InputFile = originalFile

dxfdwgCreator1.ExportDestination = NXOpen.BaseCreator.ExportDestinationOption.Teamcenter

dxfdwgCreator1.DatasetName = currentFile & strRevision

dxfdwgCreator1.OutputFileExtension = "dwg"

dxfdwgCreator1.TextFontMappingFile = "C:\Users\JUANPA~1\AppData\Local\Temp\juan1B181C15e4o5.txt"

dxfdwgCreator1.WidthFactorMode = NXOpen.DxfdwgCreator.WidthfactorMethodOptions.AutomaticCalculation

dxfdwgCreator1.CrossHatchMappingFile = "C:\Users\JUANPA~1\AppData\Local\Temp\juan1B181C15e4o6.txt"

dxfdwgCreator1.LineFontMappingFile = "C:\Users\JUANPA~1\AppData\Local\Temp\juan1B181C15e4o7.txt"

dxfdwgCreator1.LayerMask = "1-256"

dxfdwgCreator1.DrawingList = """Sheet 1"""

dxfdwgCreator1.ProcessHoldFlag = True

Dim nXObject1 As NXOpen.NXObject = Nothing
nXObject1 = dxfdwgCreator1.Commit()

theSession.DeleteUndoMark(markId3, Nothing)

dxfdwgCreator1.Destroy()

theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING")

theSession.CleanUpFacetedFacesAndEdges()

workPart.Drafting.EnterDraftingApplication()

workPart.Views.WorkView.UpdateCustomSymbols()

theSession.CleanUpFacetedFacesAndEdges()
