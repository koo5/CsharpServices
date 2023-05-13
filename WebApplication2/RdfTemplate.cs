﻿using System.Text.RegularExpressions;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using VDS.RDF;
using VDS.RDF.Parsing;
using VDS.RDF.Writing;
using VDS.RDF.Nodes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Presentation;
#if VSTO
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
#endif



#if !VSTO



#endif




namespace LodgeiT
{


    // a mapping from field to Pos
    public class FieldMap : Dictionary<INode, Pos>
    {

    }


    /*
    query / execution context, used for obtaining useful error messages
    */

    class CtxItem
    {
        public string value;
        public CtxItem(string value)
        {
            this.value = value;
        }
    }

    class GlobalCtx
    {
        public static List<CtxItem> items = new List<CtxItem>();
        public static CtxItem Add(string item)
        {
            var i = new CtxItem(item);
            items.Add(i);
            return i;
        }
        public static string PrettyString()
        {
            string result = "";
            foreach (var i in items)
            {
                result += i.value + "\n";
            }
            return result;
        }
    }

    class Ctx
    {
        List<CtxItem> local_items = new List<CtxItem>();
        public Ctx(string description)
        {
            add(description);
        }
        public Ctx(string format, params object[] args)
        {
            add(String.Format(format, args));
        }
        public void add(string context)
        {
            local_items.Add(GlobalCtx.Add(context));
        }
        public void add(string format, params object[] args)
        {
            add(String.Format(format, args));
        }
        public void pop()
        {
            var i = local_items.Last();
            if (i != GlobalCtx.items.Last())
                throw new Exception("execution context stack mismatch, this shouldn't happen");
            local_items.Remove(i);
            GlobalCtx.items.Remove(i);
        }
    }



    /// <summary>
    /// abstraction of excel cell coordinates
    /// </summary>
    public enum CellReadingResult
    {
        Ok,
        Error,
        Empty
    }

    public enum ResultSheetsHandling
    {
        InPlace,
        NewSheet,
        NewWorkbook
    }

    class UriLabelPair
    {
        public string uri;
        public string label;
        public UriLabelPair(string _uri, string _label)
        {
            uri = _uri;
            label = _label;
        }
    }

    class SheetInstanceData
    {
        public string name;
        public INode record;
        public SheetInstanceData(string name, INode record)
        {
            this.name = name;
            this.record = record;
        }
    }

    public class Pos
    {
        public int col = 'A';
        public int row = 1;
        public string Cell { get { return GetExcelColumnName() + row.ToString(); } }
        public override string ToString() { return Cell; }
        public Pos Clone() { return (Pos)MemberwiseClone(); }
        /*fixme*/
        static public int ColFromString(string s)
        {
            return s[0]; /*fixme*/
        }
        public string GetExcelColumnName()
        /* https://stackoverflow.com/a/182924 */
        {
            int dividend = col - 'A' + 1;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
    }
    public class RdfSheetEntry
    {
        public Pos _pos;
        public INode _obj;
        public RdfSheetEntry(INode obj, Pos pos)
        {
            if (pos != null)
                _pos = pos.Clone();
            _obj = obj;
        }
    }
    public class RdfTemplateInputError : Exception
    {
        public RdfTemplateInputError(string message = "") : base(message)
        {
        }
    }













    /// <summary>
    /// schema-directed UI.
    /// In future, we should probably align the templating structure with http://datashapes.org/forms.html, possibly through inferencing one from the other.
    /// 
    /// control flow / error reporting:
    /// reporting errors to user is the competence of the function that detects the error. Error is displayed and false is returned.
    /// 
    /// RdfTemplate lifecycle:
    /// construct, call one of: GenerateTemplate, ExtractSheetGroupData or DisplayResponse, and dispose
    /// 
    /// </summary>
    public class RdfTemplate// : TemplateGenerator
    {
        private INode _sheetsGroupTemplateUri;
        // the sheet currently being read or populated:
#if VSTO
        private Worksheet _sheet;
        Excel.Application _app;
#else
        private IXLWorksheet _sheet;
        XLWorkbook _app;
        public string alerts;
#endif

        private readonly bool _isFreshSheet = true;
        // This is the main graph used throughout the lifetime of RdfTemplate. It is populated either with RdfTemplates.n3, or with response.n3. response.n3 contains also the templates, because they are sent with the request. We should maybe only send the data that user fills in, but this works:
        protected Graph _g;
        // here we put core request data that can be used to construct an example sheetset from a request:
        protected Graph _rg;

        // we generate some pseudo blank nodes, unique uris. But blank nodes work too.
        protected decimal _freeBnId = 0;


#if VSTO
        public RdfTemplate(Excel.Application app)
        {
#if !DEBUG
            try
            {
#endif
            _app = app;
            Init(app);
#if !DEBUG
            }
            catch (Exception e)
            {
                MessageBox.Show("while initializing RdfTemplate: " + e.Message, "LodgeIt", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw e;
            }
#endif
        }
        public RdfTemplate(Excel.Application app, string sheetsTemplateQName)
        {
            RdfTemplate(app, u(sheetsTemplateQName));
        }
        public RdfTemplate(Excel.Application app, Uri sheetsTemplateUri)
        {
#if !DEBUG
            try
            {
#endif
            _app = app;
            Init();
            _sheetsGroupTemplateUri = _g.CreateUriNode(sheetsTemplateUri);
#if !DEBUG
            }
            catch (Exception e)
            {
                MessageBox.Show("while initializing RdfTemplate(" + sheetsTemplateUri.ToString() + "): " + e.Message, "LodgeIt", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw e;
            }
#endif
        }

#else

        public RdfTemplate(XLWorkbook app, string sheetsTemplateQName)
        {
#if !DEBUG
            try
            {
#endif
            _app = app;
            Init();
            _sheetsGroupTemplateUri = u(sheetsTemplateQName);
#if !DEBUG
            }
            catch (Exception e)
            {
                MessageBox.Show("while initializing RdfTemplate(" + sheetsTemplateUri.ToString() + "): " + e.Message, "LodgeIt", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw e;
            }
#endif
        }

#endif
        private void Init()
        {
            _g = new Graph();
            _g.NamespaceMap.AddNamespace("l", UriFactory.Create("https://rdf.lodgeit.net.au/v1/request#"));
            _g.NamespaceMap.AddNamespace("excel", UriFactory.Create("https://rdf.lodgeit.net.au/v1/excel#"));
            _g.NamespaceMap.AddNamespace("depr", UriFactory.Create("https://rdf.lodgeit.net.au/v1/calcs/depr#"));
            _g.NamespaceMap.AddNamespace("ic", UriFactory.Create("https://rdf.lodgeit.net.au/v1/calcs/ic#"));
            _g.NamespaceMap.AddNamespace("hp", UriFactory.Create("https://rdf.lodgeit.net.au/v1/calcs/hp#"));
            _g.NamespaceMap.AddNamespace("depr_ui", UriFactory.Create("https://rdf.lodgeit.net.au/v1/calcs/depr/ui#"));
            _g.NamespaceMap.AddNamespace("ic_ui", UriFactory.Create("https://rdf.lodgeit.net.au/v1/calcs/ic/ui#"));
            _g.NamespaceMap.AddNamespace("hp_ui", UriFactory.Create("https://rdf.lodgeit.net.au/v1/calcs/hp/ui#"));
            _g.NamespaceMap.AddNamespace("smsf_ui", UriFactory.Create("https://rdf.lodgeit.net.au/v1/calcs/smsf/ui#"));
            _g.NamespaceMap.AddNamespace("xsd", UriFactory.Create("http://www.w3.org/2001/XMLSchema#"));
            _g.NamespaceMap.AddNamespace("rdf", UriFactory.Create("http://www.w3.org/1999/02/22-rdf-syntax-ns#"));
            _g.NamespaceMap.AddNamespace("rdfs", UriFactory.Create("http://www.w3.org/2000/01/rdf-schema#"));

            _rg = new Graph();
            _rg.NamespaceMap.AddNamespace("rdfs", UriFactory.Create("http://www.w3.org/2000/01/rdf-schema#"));
            _rg.NamespaceMap.AddNamespace("excel", UriFactory.Create("https://rdf.lodgeit.net.au/v1/excel#"));
            _rg.NamespaceMap.AddNamespace("", UriFactory.Create("https://rdf.lodgeit.net.au/v1/excel_request#"));
            /* BaseUri is also the graph uri, for some strange reason */
            _rg.BaseUri = uu("l:request_graph");
        }

#if VSTO

        private void ErrMsg(string msg)
        {
            System.Console.WriteLine(msg);
            MessageBox.Show(msg, "LodgeiT", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
#else
        private void ErrMsg(string msg)
        {
            System.Console.WriteLine(msg);
            alerts += msg + "\n";
        }


#endif

#if VSTO

        public List<string> AvailableSheetSets(string rdf_templates)
        {
            var ctx = new Ctx("AvailableSheetSets({0})", rdf_templates);
            LoadTemplates(rdf_templates);
            List<string> result = new List<string>();
            foreach (var i in GetSubjects(u("rdf:type"), u("excel:sheet_set")))
                result.Add(i.AsValuedNode().AsString());
            ctx.pop();
            return result;
        }
        public List<UriLabelPair> ExampleSheetSets(string rdf_templates)
        {
            LoadTemplates(rdf_templates);
            List<UriLabelPair> result = new List<UriLabelPair>();
            foreach (var i in GetSubjects(u("rdf:type"), u("excel:example_sheet_set")))
            {
                if (BoolObjectWithDefault(i, u("excel:is_listed_by_hardcoding"), false))
                    continue;
                string uri = i.AsValuedNode().AsString();
                string label = GetLabels(i).First();
                UriLabelPair p = new UriLabelPair(uri, label);
                result.Add(p);
            }
            return result;
        }

        public void CreateSheetsFromExample(string rdf_templates)
        {
#if !DEBUG
            try
            {
#endif
            LoadTemplates(rdf_templates);

            foreach (INode example_sheet_info in GetListItems((_sheetsGroupTemplateUri), "excel:has_sheet_instances"))
            {
                INode sheet_decl = GetObject(example_sheet_info, "excel:sheet_instance_has_sheet_type");
                string sheet_name = GetSheetNamePrefix(sheet_decl);
                var template = GetObject(sheet_decl, "excel:root");
                var doc = GetObject(example_sheet_info, "excel:sheet_instance_has_sheet_data");
                _sheet = NewWorksheet(sheet_name, GetMultipleSheetsAllowed(sheet_decl));
                if (_sheet != null)
                {
                    WriteFirstRow(sheet_decl);
                    WriteData(template, doc);
                    _sheet.Columns.AutoFit();
                }
            }
#if !DEBUG
            }
            catch (Exception e)
            {
                ErrMsg("while CreateSheetsFromExample(" + rdf_templates + "): " + e.Message);
                throw e;
            }
#endif
        }
#endif
        private bool CreateRdfEndpointRequestFromSheetGroupData()
        {
            IEnumerable<INode> known_sheets = GetListItems(_sheetsGroupTemplateUri, "excel:sheets");
            var extracted_instances_by_sheet_type = new Dictionary<INode, IList<SheetInstanceData>>();
            if (!ExtractDataInstances(known_sheets, ref extracted_instances_by_sheet_type))
                return false;
            if (!make_sure_all_non_optional_sheets_are_present(known_sheets, extracted_instances_by_sheet_type))
                return false;
            if (!AssertRequest(extracted_instances_by_sheet_type))
                return false;
            return true;
        }

        private bool AssertRequest(Dictionary<INode, IList<SheetInstanceData>> extracted_instances_by_sheet_type)
        {
            IList<INode> all_request_sheets = new List<INode>();

            foreach (KeyValuePair<INode, IList<SheetInstanceData>> kv in extracted_instances_by_sheet_type)
            {
                INode sheet_type = kv.Key;
                var sheet_instances = kv.Value;
                if (!GetMultipleSheetsAllowed(sheet_type))
                {
                    if (sheet_instances.Count > 1)
                    {
                        ErrMsg("only one sheet of type \"" + GetSheetNamePrefix(sheet_type) + "\" (" + sheet_type.ToString() + ") is allowed.");
                        return false;
                    }
                }
                foreach (SheetInstanceData sheet_instance_data in sheet_instances)
                {
                    INode sheet_instance = _rg.CreateBlankNode();
                    Assert(_rg, sheet_instance, u(_rg, "excel:sheet_instance_has_sheet_type"), sheet_type);
                    Assert(_rg, sheet_instance, u(_rg, "excel:sheet_instance_has_sheet_name"), _g.CreateLiteralNode(sheet_instance_data.name));
                    Assert(_rg, sheet_instance, u(_rg, "excel:sheet_instance_has_sheet_data"), sheet_instance_data.record);
                    all_request_sheets.Add(sheet_instance);
                }
            }

            Assert(_rg, u(_rg, ":request"), u(_rg, "excel:has_sheet_instances"), _rg.AssertList(all_request_sheets));
            Assert(_g, u(":request"), u("l:client_version"), _g.CreateLiteralNode("2"));
            Assert(_g, u(":request"), u("l:client_git_info"), _g.CreateLiteralNode(Properties.Resources.ResourceManager.GetObject("repo_status").ToString().Replace("\n", Environment.NewLine)));
            return true;
        }

        private bool ExtractDataInstances(IEnumerable<INode> known_sheets, ref Dictionary<INode, IList<SheetInstanceData>> extracted_instances_by_sheet_type)
        {
            foreach (Excel.Worksheet sheet in _app.Worksheets)
            {
                _sheet = sheet;

                // ignore any sheet that does not have the type header
                if (GetCellValue(new Pos { col = 'A', row = 1 }).Trim().ToLower() != "sheet type:")
                    continue;

                string sheet_type_uri_string = GetCellValue(new Pos { col = 'B', row = 1 }).Trim();
                INode sheet_type_uri = _g.CreateUriNode(new Uri(sheet_type_uri_string));
                if (!known_sheets.Contains(sheet_type_uri))
                {
                    ErrMsg("unknown sheet type: " + sheet_type_uri_string + ", ignoring.");
                    return false;
                }
                INode record_instance = null;
                INode sheet_template = GetObject(sheet_type_uri, u("excel:root"));
                if (!ExtractRecordByTemplate(sheet_template, ref record_instance))
                    return false;
                //Assert(_g, record_instance, u("excel:has_sheet_name"), AssertValue(_g, _g.CreateLiteralNode(sheet.Name)));
                Assert(_g, record_instance, u("excel:sheet_type"), sheet_type_uri);
                if (!extracted_instances_by_sheet_type.ContainsKey(sheet_type_uri))
                    extracted_instances_by_sheet_type[sheet_type_uri] = new List<SheetInstanceData>();
                extracted_instances_by_sheet_type[sheet_type_uri].Add(new SheetInstanceData(_sheet.Name, record_instance));
            }
            return true;
        }
        private bool make_sure_all_non_optional_sheets_are_present(IEnumerable<INode> known_sheets, Dictionary<INode, IList<SheetInstanceData>> extracted_instances_by_sheet_type)
        {
            foreach (INode known_sheet in known_sheets)
            {
                if (BoolObjectWithDefault(known_sheet, u("excel:optional"), false))
                    continue;
                if (extracted_instances_by_sheet_type.ContainsKey(known_sheet))
                    continue;
                string msg = "sheet \"" + GetSheetNamePrefix(known_sheet) + "\" (" + known_sheet.ToString() + ") not found.";
                ErrMsg(msg);
                return false;
            }
            return true;
        }
        public bool ExtractSheetGroupData(string UpdatedRdfTemplates = "")
        {
            LoadTemplates(UpdatedRdfTemplates);
            try
            {
                return this.CreateRdfEndpointRequestFromSheetGroupData();
            }
            catch (RdfTemplateInputError)
            {
                return false;
            }
        }

        private bool GetMultipleSheetsAllowed(INode sheet_decl)
        {
            return BoolObjectWithDefault(sheet_decl, u("excel:multiple_sheets_allowed"), false);
        }

        private string GetSheetNamePrefix(INode sheet)
        {
            var n = GetObjects(sheet, "excel:name_prefix");
            if (!n.Any())
            {
                IUriNode r = (IUriNode)GetObject(sheet, "excel:root");
                return UriFragment(r);
            }
            return n.First().AsValuedNode().AsString();
        }

        public void LoadRequestSheets(StreamReader data)
        {
            /* todo. 
            what namespace is request in? 
            what graph?
            */
        }

        public void LoadResultSheets(StreamReader data, ResultSheetsHandling result_sheet_handling)
        {
            LoadRdf(data);
            var instances = GetSubjects(u("excel:is_result_sheet"), true.ToLiteral(_g));
            DisplaySheets(instances, result_sheet_handling);
        }

        public void DisplaySheets(IEnumerable<INode> sheets, ResultSheetsHandling result_sheet_handling)
        {
            if (sheets.Count() != 0 && result_sheet_handling == ResultSheetsHandling.NewWorkbook)
            {
                Workbook wb = _app.Workbooks.Add();
                //todo: bring it to the front or something?
            }
            foreach (var sheet_instance in sheets)
            {
                string sheet_name = SanitizeSheetName(((ILiteralNode)GetObject(sheet_instance, u("excel:sheet_instance_has_sheet_name"))).Value);
                var sheet_type = GetObject(sheet_instance, "excel:sheet_instance_has_sheet_type");
                var doc = GetObject(sheet_instance, "excel:sheet_instance_has_sheet_data");
                var template = GetObject(sheet_type, "excel:root");

                if (result_sheet_handling == ResultSheetsHandling.InPlace)
                    _sheet = SheetByName(sheet_name);
                else
                    _sheet = NewWorksheet(sheet_name, true);

                WriteFirstRow(sheet_type);
                WriteData(template, doc);

                _sheet.Columns.AutoFit();
            }
        }

        public string SanitizeSheetName(string old)
        {
            // https://stackoverflow.com/questions/451452/valid-characters-for-excel-sheet-names/451488
            const string invalidCharsRegex = @"[/\\*'?[\]:]+";
            const int maxLength = 31;

            string safeName = Regex.Replace(old, invalidCharsRegex, " ")
                                    .Replace("  ", " ")
                                    .Trim();

            if (string.IsNullOrEmpty(safeName))
            {
                safeName = "sheet";   // cannot be empty
            }
            else if (safeName.Length > maxLength)
            {
                safeName = safeName.Substring(0, maxLength);
            }
            return safeName;
        }

        public void WriteFirstRow(INode sheet_decl)
        {
            WriteString(new Pos { col = 'A', row = 1 }, "sheet type:");
            //_sheet.Range["A1"].AddComment("blablabl\nablablabla");
            WriteString(new Pos { col = 'B', row = 1 }, sheet_decl.ToString());
        }

        public bool ExtractRecordByTemplate(INode template, ref INode individual)
        {
            var map = new FieldMap();
            var unknown_fields = new List<INode>();
            if (!HeadersMapping(template, ref map, ref unknown_fields))
                return false;
            if (!IsMulti(template))
            {
                if (!ExtractRecord(template, map, 1, ref individual, true))
                    return false;
            }
            else
            {
                int item_offset = 1;
                var numEmptyRows = 0;
                var individuals = new List<INode>();
                do
                {
                    INode item = null;
                    if (ExtractRecord(template, map, item_offset, ref item, false))
                    {
                        if (item != null)
                        {
                            individuals.Add(item);
                            numEmptyRows = 0;
                        }
                        else
                            numEmptyRows++;
                    }
                    else
                        numEmptyRows++;
                    item_offset++;
                } while (numEmptyRows < 5);
                var rdf_list = _rg.AssertList(individuals);
                individual = Bn(_rg, "list");
                Assert(_rg, individual, u("rdf:value"), rdf_list);
            }
            Assert(_rg, individual, u("excel:template"), template);
            Assert(_g, individual, u("excel:has_sheet_name"), _sheet.Name.ToLiteral(_g));
            if (unknown_fields.Count > 0)
                Assert(_rg, individual, u("excel:has_unknown_fields"), _rg.AssertList(unknown_fields));
            return true;
        }

        protected FieldMap GetRecordCellPositions(INode template, FieldMap map, int item_offset)
        {
            bool is_horiz = GetIsHorizontal(template);
            var cell_positions = new FieldMap();
            foreach (KeyValuePair<INode, Pos> mapping in map)
            {
                Pos pos = mapping.Value.Clone();
                if (is_horiz)
                    pos.col += item_offset;
                else
                    pos.row += item_offset;
                cell_positions[mapping.Key] = pos;
            }
            return cell_positions;
        }

        protected void AssertRecordPos(INode template, FieldMap cell_positions, ref INode record)
        {
            if (IsMulti(template))
            {
                if (cell_positions.Keys.Count > 0)
                    AssertPosFlat(record, cell_positions.First().Value);
            }
            else
                AssertPosFlat(record, GetPos(template));
        }

        protected bool ReadSubTemplates(INode template, ref Dictionary<INode, RdfSheetEntry> values)
        {
            INode obj = null;
            foreach (var field in GetFields(template))
            {
                //string qdesc1 = null;
                var field_templates = GetObjects(field, u("excel:template"));
                if (field_templates.Any())
                {
                    INode field_template = null;
                    try
                    {
                        field_template = One(field_templates);
                    }
                    catch (Exception ex)
                    {
                        ErrMsg(field.ToString() + " excel:template: " + ex.Message);
                        throw ex;
                    }
                    if (!ExtractRecordByTemplate(field_template, ref obj))
                        return false;
                    if (obj != null)
                        values[GetPropertyUri(field)] = new RdfSheetEntry(obj, null);
                }
            }
            return true;
        }


        protected bool ExtractRecord(INode template, FieldMap map, int item_offset, ref INode record, bool isRequired)
        {
            var cls = GetClass(template);
            var values = new Dictionary<INode, RdfSheetEntry>();
            var cell_positions = GetRecordCellPositions(template, map, item_offset);
            if (!ReadCellValues(cell_positions, ref values))
                return false;
            if (!ReadSubTemplates(template, ref values))
                return false;
            if (!values.Any())
            {
                if (isRequired)
                {
                    String msg = "no values detected in template " + GetTitle(template) + ". Expected:\n";
                    foreach (KeyValuePair<INode, Pos> mapping in cell_positions)
                        msg += FieldTitles(mapping.Key).First() + " at " + mapping.Value.Cell + "\n";
                    ErrMsg(msg);
                }
                return false;
            }
            record = Bn(_rg, "record");
            Assert(_rg, record, u("rdf:type"), cls);
            AssertRecordPos(template, cell_positions, ref record);
            Assert(_g, record, u("excel:has_sheet_name"), _sheet.Name.ToLiteral(_g));
            foreach (KeyValuePair<INode, RdfSheetEntry> entry in values)
            {
                Assert(_rg, record, entry.Key, entry.Value._obj);
                if (entry.Value._pos != null)
                    AssertPosFlat(entry.Value._obj, entry.Value._pos);
                Assert(_g, entry.Value._obj, u("excel:has_sheet_name"), _sheet.Name.ToLiteral(_g));
            }
            return true;
        }

        protected bool ReadCellValues(FieldMap cell_positions, ref Dictionary<INode, RdfSheetEntry> values)
        {
            foreach (KeyValuePair<INode, Pos> mapping in cell_positions)
            {
                INode field = mapping.Key;
                INode obj = null;
                Pos pos = cell_positions[field];
                IEnumerable<INode> types = GetObjects(field, "excel:type");

                if (!ReadCellAsType(pos, field, types, ref obj))
                    return false; //either error or possibly end of sheet

                if (obj != null)
                    values[GetPropertyUri(field)] = new RdfSheetEntry(AssertValue(_rg, obj), pos);
                else if (!BoolObjectWithDefault(field, u("excel:optional"), true))
                {
                    ErrMsg("missing required field in " + _sheet.Name + " at " + pos.Cell);
                    return false;
                }
            }
            return true;
        }

        bool ReadCellAsType(Pos pos, INode field, IEnumerable<INode> types, ref INode obj)
        {
            INode type;
            if (!types.Any())
                type = u("xsd:string");
            else
                type = types.First();
            if (type.Equals(u("xsd:decimal")))
            {
                if (!ReadOptionalDecimal(pos, ref obj))
                    return false;
            }
            else if (type.Equals(u("xsd:integer")))
            {
                if (!ReadOptionalInt(pos, ref obj))
                    return false;
            }
            else if (type.Equals(u("xsd:dateTime")))
            {
                DateTime contents = ExporttoXMLBase.GetCellAsDate(_sheet, pos.Cell);
                if (contents != DateTime.MinValue)
                    obj = contents.Date.ToLiteral(_g);
                /*else
                    if (CellStringContents(pos) != "")
                    {
                        ErrMsg("error reading date in \"" + _sheet.Name + "\" at " + pos.Cell);
                        return false;
                    }*/
                else
                {
                    string contents_str = GetCellValue(pos);
                    if (contents_str != "")
                        obj = contents_str.ToLiteral(_g);
                }
            }
            else if (type.Equals(u("excel:uri")))
            {
                obj = GetUriObj(pos, field);
            }
            else if (type.Equals(u("xsd:string")))
            {
                string contents = null;
                CellReadingResult status = GetCellValueAsString(pos, ref contents);
                if (status == CellReadingResult.Ok)
                    obj = contents.ToLiteral(_g);
                else if (status == CellReadingResult.Error) // maybe added for end of sheet data, but i'm not sure it makes sense
                    return false;
            }
            else
                throw new Exception("RDF template error: excel:type not recognized: " + type.ToString());
            return true;
        }


        /*
        todo
        public CellReadingResult GetCellValueAsInteger(Pos pos, ref int result)
        {
            Range rng = _sheet.get_Range(pos.Cell, pos.Cell);
            if (rng.Value2 is Int32)
            {
                ErrMsg("error in " + _sheet.Name + " " + pos.Cell);
                return CellReadingResult.Error;
            }
            string sValue = null;
            if (rng.Value2 != null)
                sValue = Convert.ToString(rng.Value2);
            else
                sValue = Convert.ToString(rng.Text);
            if (sValue == null)
                return CellReadingResult.Empty;
            if (sValue.Length == 0)
                return CellReadingResult.Empty;
            string sVal = sValue.Trim('$');
            if (int.TryParse(sVal, out result))
                return CellReadingResult.Ok;
            else
                return CellReadingResult.Error;
        }
        */

        public string CellStringContents(Pos pos)
        {
            string contents = "";
            CellReadingResult status = GetCellValueAsString(pos, ref contents);
            return contents;

        }
        public bool ReadOptionalInt(Pos pos, ref INode obj)
        /*
        return true if cell is empty
    set obj and return true on successful parse	
    show messagebox and return false on parse error.
        */
        {
            string sValue = GetCellValue(pos).Trim();
            //sValue = sValue.Trim('$');
            int result;
            if (sValue == "")
                return true;
            else
            {
                if (!int.TryParse(sValue, out result))
                {
                    ErrMsg("error reading integer in " + _sheet.Name + " at " + pos.Cell);
                    return false;
                }
            }
            obj = result.ToLiteral(_g);
            return true;
        }
        public bool ReadOptionalDecimal(Pos pos, ref INode obj)
        /*
        return true if cell is empty
        set obj and return true on successful parse	
        show messagebox and return false on parse error.
        */
        {
            string sValue = GetCellValue(pos).Trim();
            sValue = sValue.Trim('$');
            if (sValue == "")
                return true;

            Excel.Range rng = _sheet.get_Range(pos.Cell, pos.Cell);
            decimal result = 0;

            if (rng.Value != null)
            {
                try
                {
                    result = ((IConvertible)rng.Value).ToDecimal(null);
                    obj = Math.Round(result, 6).ToLiteral(_g);
                    return true;
                }
                catch (System.FormatException e)
                {
                    ErrMsg("error reading decimal in " + _sheet.Name + " at " + pos.Cell + ", got: \"" + sValue + "\", error: " + e.Message);
                    throw new RdfTemplateInputError();
                }
            }

            if (!decimal.TryParse(sValue, out result))
            {
                ErrMsg("error reading decimal in " + _sheet.Name + " at " + pos.Cell + ", got: \"" + sValue + "\"");
                throw new RdfTemplateInputError();
            }
            obj = result.ToLiteral(_g);
            return true;
        }

        public CellReadingResult GetCellValueAsString(Pos pos, ref string result)
        {
            Range rng = _sheet.get_Range(pos.Cell, pos.Cell);
            if (rng.Value2 is Int32)
            {
                ErrMsg("error in " + _sheet.Name + " " + pos.Cell);
                return CellReadingResult.Error;
            }
            string sValue = null;
            if (rng.Value2 != null)
                sValue = Convert.ToString(rng.Value2);
            else
                sValue = Convert.ToString(rng.Text);
            if (sValue == null)
                return CellReadingResult.Empty;
            sValue = sValue.Trim();
            if (sValue.Length == 0)
                return CellReadingResult.Empty;
            result = sValue/**.Trim('$')*/;
            return CellReadingResult.Ok;
        }

        public void Assert(IGraph g, INode s, INode p, INode o)
        {
            g.Assert(new Triple(Tools.CopyNode(s, g), Tools.CopyNode(p, g), Tools.CopyNode(o, g)));
        }
        public INode AssertValue(IGraph g, INode obj)
        {
            var value = Bn(g, "value");
            //g.Assert(new Triple(value, u("rdf:type"), u("l:value")));
            Assert(g, value, u("rdf:value"), obj);
            //g.Assert(new Triple(value, _g.CreateUriNode("l:unit"), _g.CreateVariableNode("u")));
            return value;
        }
        public INode GetValue(INode s)
        {
            return GetObject(s, u("rdf:value"));
        }
        private INode GetUriObj(Pos pos, INode field)
        {

            string contents = GetCellValue(pos);
            if (contents.Equals(""))
                return null;
            var prop = GetPropertyUri(field);
            var range = GetRange(prop);
            foreach (var option in GetSubjects(u("rdf:type"), range))
            {
                var option_labels = GetLabels(option);
                foreach (var label in option_labels)
                {
                    if (contents.Equals(label))
                        return option;
                }
            }
            try
            {
                return u(contents);
            }
            catch (VDS.RDF.RdfException)
            {
                string msg = "invalid URI: " + contents;
                ErrMsg(msg);
                throw;
            }
        }

        /*find a header of each field, produce a mapping from field to Pos*/
        protected bool HeadersMapping(INode template, ref FieldMap result, ref List<INode> unknown_fields)
        {
            var pos = GetPos(template);
            bool is_horiz = GetIsHorizontal(template);
            // skip the section title
            pos.row += 1;
            int nested_templates_count = 0;
            var fields = GetFields(template);
            if (fields.Count() <= 0)
                throw new Exception("invalid template: fields.Count() <= 0");


            List<string> expected_header_items = new List<string>();
            {
                foreach (var field in fields)
                {
                    var field_template = GetObjects(field, "excel:template");
                    if (field_template.Any())
                    {
                        nested_templates_count++;
                        continue;
                    }
                    var titles = FieldTitles(field);
                    expected_header_items.Add(titles.First());
                }
            }


            foreach (int item in Enumerable.Range(0, 100))//fixme
            {
                Pos pos2 = pos.Clone();
                if (is_horiz)
                    pos2.row += item;
                else
                {
                    pos2.col += item;
                    if (item > 25)//fixme
                        break;
                }
                string sKey = GetCellValue(pos2).ToLower();
                if (sKey == "") continue;
                bool header_belongs_to_matching_field_declaration = false;
                foreach (var field in fields)
                {
                    var field_template = GetObjects(field, "excel:template");
                    if (field_template.Any())
                    {
                        nested_templates_count++;
                        continue;
                    }
                    var titles = FieldTitles(field);
                    if (titles.Contains(sKey, StringComparer.OrdinalIgnoreCase))
                    {
                        result[field] = pos2;
                        header_belongs_to_matching_field_declaration = true;
                        break;
                    }
                }
                if (!header_belongs_to_matching_field_declaration)
                {
                    if (BoolObjectWithDefault(template, u("excel:grab_unknown_fields"), false))
                    {
                        INode field = PseudoBn(_rg, "field");
                        unknown_fields.Add(field);
                        INode header_cell_value = null;
                        if (!ReadCellAsType(
                                pos2,
                                u("excel:has_header_cell_value"), /*kinda dummy, fixme*/
                                GetObjects(template, u("excel:unknown_fields_type")),
                                ref header_cell_value
                            ))
                            return false;
                        Assert(_g,
                            field,
                            u("excel:has_header_cell_value"),
                            header_cell_value);
                        result[field] = pos2;
                    }
                }
            }

            if (nested_templates_count == 0 && !result.Any())
            {
                string msg = "No headers found, for template \"" + GetTitle(template) + "\", sheet \"" + _sheet.Name + "\".\nexpected: ";
                foreach (string i in expected_header_items)
                    msg += i + ", ";
                msg += "starting from " + pos.Cell;
                ErrMsg(msg);
                return false;
            }
            return true;
        }

        void WriteData(INode template, INode doc)
        {
            var pos = GetPos(template);
            if (_isFreshSheet)
                PopulateHeader(pos.Clone(), template);
            bool is_horiz = GetIsHorizontal(template);
            // go one row down from title
            pos.row++;
            if (is_horiz)
                pos.col++; // go one col right from header
            else
                pos.row++;
            int numCellsToMakeDropdownsOn = 1;
            if (!IsMulti(template))
                PopulateData(pos.Clone(), template, doc);
            else if (doc != null)
                numCellsToMakeDropdownsOn += PopulateRows(pos.Clone(), template, doc) + 10;
            MakeDropdowns(pos.Clone(), template, numCellsToMakeDropdownsOn);
        }
        protected int PopulateRows(Pos pos, INode template, INode doc)
        {
            int rows = 0;
            var rdf_list = GetObject(doc, "rdf:value");
            foreach (var item in _g.GetListItems(rdf_list))
            {
                rows++;
                PopulateData(pos.Clone(), template, item);
                pos.row++;
            }
            return rows;
        }
        void MakeDropdowns(Pos pos, INode template, int numCellsToMakeDropdownsOn)
        {
            foreach (var field in GetFields(template))
            {
                if (GetObjects(field, "excel:template").Any())
                    continue;
                if (GetObjects(field, "excel:type").Contains(u("excel:uri")))
                    MakeDropdownsForField(pos.Clone(), field, template, numCellsToMakeDropdownsOn);
                if (GetIsHorizontal(template))
                    pos.row++;
                else
                    pos.col++;
            }
        }
        void MakeDropdownsForField(Pos pos, INode field, INode template, int numCellsToMakeDropdownsOn)
        {
            var prop = GetPropertyUri(field);
            var validation_string = GetValidationString(prop);
            foreach (int i in Enumerable.Range(0, numCellsToMakeDropdownsOn))
            {
                MakeDropdown(pos, validation_string);
                if (GetIsHorizontal(template))
                    pos.col++;
                else
                    pos.row++;
            }
        }
        void MakeDropdown(Pos pos, string validation_string)
        {
            Debug.WriteLine(_sheet.Name + " " + pos.Cell + " Validation.Add:" + validation_string);
            _sheet.Range[pos.Cell].Validation.Add(
                XlDVType.xlValidateList,
                XlDVAlertStyle.xlValidAlertInformation,
                XlFormatConditionOperator.xlBetween,
                validation_string);
        }
        string GetValidationString(INode prop)
        {
            var range = GetRange(prop);
            var options = GetSubjects(u("rdf:type"), range);
            List<string> labels = new List<string>();
            foreach (INode option in options)
                labels.Add(GetLabels(option).First());
            return string.Join(",", labels);
        }
        void PopulateData(Pos pos, INode template, INode doc)
        {
            foreach (var field in GetFields(template))
            {
                var prop = GetPropertyUri(field);
                var field_template = GetObjects(field, "excel:template");
                INode field_value = null;
                if (doc != null)
                {
                    var field_values = GetObjects(doc, prop);
                    if (field_values.Any())
                        field_value = One(field_values);
                }
                if (!field_template.Any())
                {
                    if (field_value != null)
                        WriteValue(pos, field_value);
                    if (GetIsHorizontal(template))
                        pos.row++;
                    else
                        pos.col++;
                }
                else
                    WriteData(field_template.First(), field_value);
            }
        }

        void WriteValue(Pos pos, INode field_value)
        {
            INode value_value = GetObject(field_value, "rdf:value");
            if (value_value == null)
                throw new Exception("expected rdf:value, got: " + field_value.ToString() + " with no rdf:value property");
            if (value_value.NodeType == NodeType.Literal)
            {
                IValuedNode literal = value_value.AsValuedNode();
                string xsd_type = literal.EffectiveType;
                if (xsd_type == "http://www.w3.org/2001/XMLSchema#dateTime")
                    WriteDate(pos, literal.AsDateTime());
                else if (xsd_type == "http://www.w3.org/2001/XMLSchema#integer")
                    WriteInteger(pos, literal.AsInteger());
                else if (xsd_type == "http://www.w3.org/2001/XMLSchema#decimal" || xsd_type == "http://www.w3.org/2001/XMLSchema#double")
                    WriteDecimal(pos, literal.AsDecimal());
                else
                    WriteString(pos, literal.AsString());
                INode format = MaybeGetObject(field_value, u("excel:suggested_display_format"));
                IValuedNode format2 = format.AsValuedNode();
                if (format2 != null)
                    SetCellFormat(pos, format2.AsString());
            }
            else
                // it's an URI
                WriteString(pos, GetLabels(value_value).First());
        }

        void WriteDate(Pos pos, DateTime dt)
        {
            var rng = _sheet.Range[pos.Cell];
            //rng.NumberFormat = "dd/mm/yyyy";
            rng.Value = dt;
            Debug.WriteLine(_sheet.Name + " " + pos.Cell + " WriteDate:" + dt);
        }
        void WriteDecimal(Pos pos, Decimal d)
        {
            var rng = _sheet.Range[pos.Cell];
            rng.NumberFormat = "0.00";
            rng.Value = d;
            Debug.WriteLine(_sheet.Name + " " + pos.Cell + " WriteDecimal:" + d);
        }
        void WriteInteger(Pos pos, long i)
        {
            var rng = _sheet.Range[pos.Cell];
            rng.Value = i;
            Debug.WriteLine(_sheet.Name + " " + pos.Cell + " WriteInteger:" + i);
        }
        void WriteString(Pos pos, String s)
        {
            var rng = _sheet.Range[pos.Cell];
            rng.Value = s;
            Debug.WriteLine(_sheet.Name + " " + pos.Cell + " WriteString:" + s);
        }

        void SetCellFormat(Pos pos, String format)
        {
            var rng = _sheet.Range[pos.Cell];
            rng.NumberFormat = format;
        }

        void PopulateHeader(Pos pos, INode template)
        {
            var title = GetTitle(template);
            bool is_horiz = GetIsHorizontal(template);
            AddBoldValueBorder(_sheet, pos.Cell, title);
            pos.row++;
            foreach (var field in GetFields(template))
            {
                if (!GetObjects(field, "excel:template").Any())
                {
                    string cell_title = FieldTitles(field).First();
                    AddBoldValueBorder(_sheet, pos.Cell, cell_title);
                    var comments = GetObjects(field, u("excel:comment"));
                    if (comments.Any())
                    {
                        pos.col += 2;
                        WriteString(pos, "(" + comments.First().AsValuedNode().AsString() + ")");
                        pos.col -= 2;
                    }
                }
                if (is_horiz)
                    pos.row++;
                else
                    pos.col++;

            }
        }

        protected Pos GetPos(INode subject)
        {
            var uri = MaybeGetObject(subject, u("excel:position"));
            if (uri == null)
                return new Pos { col = 'A', row = 3 };
            return new Pos
            {
                col = Pos.ColFromString(One(GetObjects(uri, "excel:col")).AsValuedNode().AsString()),
                row = int.Parse(One(GetObjects(uri, "excel:row")).AsValuedNode().AsString())
            };
        }
        protected void AssertPosFlat(INode subject, Pos pos)
        {
            Assert(_g, subject, u("excel:col"), pos.GetExcelColumnName().ToLiteral(_g));
            Assert(_g, subject, u("excel:row"), pos.row.ToString().ToLiteral(_g));
        }

        protected IEnumerable<INode> GetFields(INode template)
        {
            return GetListItems(template, u("excel:fields"));
        }

        protected INode GetPropertyUri(INode field)
        {
            var prop1 = GetObjects(field, u("excel:property"));
            INode property;
            if (!prop1.Any())
                property = field;
            else
                property = One(prop1);
            return property;
        }

        //
        // Summary:
        //     We have two types of arrangements of data in sheets: For single entity, for example bank details, we use a vertical header, and for multiple entities, for example bank statement, a horizontal header
        protected bool IsMulti(INode template)
        {
            return GetObject(template, "excel:cardinality").Equals(u("excel:multi"));
        }
        protected INode GetClass(INode template)
        {
            return GetObject(template, "excel:class");
        }
        protected INode GetRange(INode prop)
        {
            return GetObject(prop, "rdfs:range");
        }
        protected IEnumerable<INode> GetListItems(INode subject, string predicate)
        {
            return GetListItems(subject, u(predicate));
        }
        protected IEnumerable<INode> GetListItems(INode subject, INode predicate)
        {
            return _g.GetListItems(One(GetObjects(subject, predicate)));
        }

        protected IEnumerable<INode> GetSubjects(INode predicate, INode obj)
        {
            foreach (Triple triple in _g.GetTriplesWithPredicateObject(predicate, obj))
                yield return triple.Subject;
        }
        protected IEnumerable<INode> GetObjects(INode subject, INode predicate)
        {
            //desc = "GetObjects(" + subject.ToString() + ", " + predicate.ToString() + ")";
            foreach (Triple triple in _g.GetTriplesWithSubjectPredicate(subject, predicate))
                yield return triple.Object;
        }
        protected IEnumerable<INode> GetObjects(INode subject, string predicate)
        {
            return GetObjects(subject, u(predicate));
        }
        protected IUriNode u(string uri)
        {
            return u(_g, uri);
        }
        protected IUriNode u(IGraph g, string uri)
        {
            return g.CreateUriNode(uri);
        }
        protected Uri uu(string uri)
        {
            return u(uri).Uri;
        }
        protected INode GetObject(INode subject, INode predicate)
        {
            return One(GetObjects(subject, predicate));
        }
        protected INode GetObject(INode subject, string predicate)
        {
            return GetObject(subject, u(predicate));
        }
        protected INode MaybeGetObject(INode subject, INode predicate)
        {
            var x = GetObjects(subject, predicate);
            if (x.Any())
                return x.First();
            return null;
        }
        protected INode One(IEnumerable<INode> e)
        {
            /*
            todo: 
            try
            {
                One(...)
            }
            catch (UnexpectedDataFormat e)
            {
                Throw UnexpectedDataFormat("when ...: " + e.Message);
            }

            [....]
            throw UnexpectedDataFormat("expected one item but got none")
            */
            if (e.Count() == 0)
                throw new Exception("RDF data error: expected one item but got none");
            if (e.Count() > 1)
                throw new Exception("RDF data error: expected one item but got multiple");
            return e.First();
        }

        bool BoolObjectWithDefault(INode subj, INode pred, bool defa)
        {
            INode node = MaybeGetObject(subj, pred);
            if (node == null)
                return defa;
            return node.AsValuedNode().AsSafeBoolean();
        }
        protected string GetCellValue(Pos pos)
        {
            return ExporttoXMLBase.GetCellValue(_sheet, pos.Cell);
        }

        protected IEnumerable<string> FieldTitles(INode field)
        {
            /* currently, the name of the field, displayed in the header, is simply the uri of the property:
             * 	:fields (
             * 	    :a 
             * 	    [:property :b])
             * 	declares that the template has two fields, :a and :b, and the expanded forms of :a and :b will be displayed.
             * 	Not sure if we eventually want to take this string from rdfs:label of the rdf:Property, from a custom property of the excel:field object, or from both.
             * 	*/
            return GetLabels(GetPropertyUri(field));
            //return new List<string>() { GetPropertyUri(field).ToString() };
        }
        protected INode PseudoBn(IGraph g, string id_base)
        {
            return u(":bn_" + id_base + (_freeBnId++).ToString());
        }
        protected INode Bn(IGraph g, string id_base)
        {
            //return u(":bn_" + id_base + (_freeBnId++).ToString());
            return g.CreateBlankNode();
        }
#if VSTO
        // isMulti = true:
        //	create, or create with a different name, if a sheet with name sheet_name already exists
        // isMulti = false:	
        //  create sheet with name sheet_name, or fail
        private Worksheet NewWorksheet(string sheet_name, bool isMulti)
        {
            if (!isMulti && SheetByName(sheet_name) != null)
            {
                ErrMsg("sheet with that name already exists: " + sheet_name);
                return null;
            }
            Worksheet worksheet = _app.Sheets.Add();
            worksheet.Name = GetUniqueName(_app.Sheets, sheet_name);
            return worksheet;
        }
        private string GetUniqueName(Excel.Sheets sheets, string prefix)
        {
            List<string> names = new List<string>();
            foreach (Excel.Worksheet sheet in sheets)
                names.Add(sheet.Name.ToLower());
            return GetUniqueName(names, prefix);
        }
        private string GetUniqueName(List<string> names, string prefix)
        {
            int counter = 0;
            string temp = prefix.ToLower();
            while (names.Contains(temp))
            {
                counter++;
                temp = prefix.ToLower() + "_" + counter.ToString();
            }

            if (counter > 0)
                prefix = prefix + "_" + counter.ToString();

            return prefix;
        }
#endif

        private bool GetIsHorizontal(INode template)
        {
            if (!IsMulti(template))
                return true;
            return BoolObjectWithDefault(template, u("excel:is_horizontal"), false);
        }
        public string Serialize()
        {
            //Notation3Writer w = new Notation3Writer();
            TriGWriter w = new TriGWriter();
            w.HighSpeedModePermitted = false;
            w.CompressionLevel = WriterCompressionLevel.High;
            w.PrettyPrintMode = true;

            return VDS.RDF.Writing.StringWriter.Write(GraphsTripleStore(), w);
        }
        public void SerializeToFile(string fn)
        {

            GZippedTriGWriter w = new GZippedTriGWriter();
            //w.CompressionLevel = VDS.RDF.Writing.WriterCompressionLevel.High;
            ITripleStore ts = GraphsTripleStore();
            w.Save(ts, fn);

            //NQuadsWriter nqw = new NQuadsWriter();
            //nqw.Save(ts, fn + ".nq");
        }

        //public static string SerializeGraph(IGraph g)
        //{
        //	/*#if DEBUG
        //				SaveAsRdfXml(g);
        //	#endif*/
        //	Notation3Writer w = new Notation3Writer();
        //	return VDS.RDF.Writing.StringWriter.Write(g, w);
        //}

        public ITripleStore GraphsTripleStore()
        {
            ITripleStore ts = new TripleStore();
            ts.Add(_g);
            ts.Add(_rg);
            return ts;
        }

#if DEBUG
        protected static void SaveAsRdfXml(IGraph g)
        {
            /*RdfXmlWriter rdfXmlWriter = new RdfXmlWriter();
            rdfXmlWriter.Save(g, "j:\\Example.rdf.xml");*/
        }
#endif
        protected void LoadRdf(StreamReader data)
        {
#if !DEBUG
            try
            {
#endif
            var parser = new Notation3Parser();
            parser.Load(_g, data);
#if !DEBUG
            }
            catch (Exception ex)
            {
                ErrMsg("Error: " + ex.Message);
                throw ex;
            }
#endif
        }
        protected void LoadTemplates(string UpdatedRdfTemplates)
        {
#if !DEBUG
            try
            {
#endif
            StreamReader reader;
            if (UpdatedRdfTemplates != null)
                /*these live-updated templates is probably the only source we should support */
                reader = new StreamReader(new MemoryStream(Encoding.UTF8.GetBytes(UpdatedRdfTemplates)));
            else
#if JINDRICH_DEBUG
                reader = new StreamReader(File.OpenRead(@"C:\Users\kokok\source\repos\LodgeITSmart\LodgeiTSmart\LodgeiTSmart\Resources\RdfTemplates.n3"));
#else
#if VSTO
                reader = new StreamReader(new MemoryStream((byte[])Properties.Resources.ResourceManager.GetObject("RdfTemplates")));
#else
                reader = new StreamReader(File.OpenRead(Environment.GetEnvironmentVariable("RDF_TEMPLATES_N3")));
#endif
#endif
            LoadRdf(reader);
#if !DEBUG
            }
            catch (Exception ex)
            {
                ErrMsg("Error: " + ex.Message);
                throw ex;
            }
#endif
        }
#if VSTO
        protected Worksheet SheetByName(string name)
        {
            foreach (Excel.Worksheet sheet in _app.Worksheets)
                if (name.Trim().ToLower() == sheet.Name.Trim().ToLower())
                    return sheet;
            return null;
        }

#endif
        protected IEnumerable<string> GetLabels(INode node)
        {
            var labels = new List<string>();
            foreach (var l in GetObjects(node, u("rdfs:label")))
                labels.Add(l.AsValuedNode().AsString());
            char[] sep = { '#' };
            labels.Add(node.ToString().Split(sep).Last());
            return labels;
        }
        protected string GetTitle(INode template0)
        {
            IUriNode template = (IUriNode)template0;
            var t = MaybeGetObject(template, u("excel:title"));
            if (t != null)
                return t.AsValuedNode().AsString();
            else
                return UriFragment(template);
        }
        protected string UriFragment(INode uri)
        {
            return ((IUriNode)uri).Uri.Fragment.Substring(1);
        }
    }

}



/*

    todo: what if a column is present twice?

*/
/*var s = wb.Worksheet("xxx");
s.Cell(1, "A").Value*/