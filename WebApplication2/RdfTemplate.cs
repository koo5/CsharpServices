using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;
using VDS.RDF;
using VDS.RDF.Parsing;
using VDS.RDF.Writing;
using VDS.RDF.Nodes;
//using Lucene.Net.Diagnostics;

#if !OOXML

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

#else

using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Presentation;
using ClosedXML.Excel;

#endif



namespace LodgeiT
{
    /*
    execution context, used for obtaining useful error messages
    */
    public class C
    {
#if !OOXML
        // this is really a matter of C# or dotnet version, but anyway
        public WeakReference parent;
#else
        public WeakReference<C> parent;
#endif
        public static C root;
        public static C current_context;
        public string value;
        public List<C> items = new List<C>();


        public C(string value)
        {
            this.value = value;
            if (C.current_context != null)
            {
#if !OOXML
                parent = new WeakReference(C.current_context);
#else
                parent = new WeakReference<C>(C.current_context);
#endif
                C.current_context.items.Add(this);
            }
            else
            {
                C.current_context = this;
                Debug.Assert(C.root == null);
                C.root = this;
            }
        }

        public void SetCurrent(C c)
        {
            current_context = c;
        }

        public string PrettyString(int indent = 0)
        {
            string result = String.Concat(Enumerable.Repeat("--", indent)) + value;
            if (items.Count > 0)
                result += ":";
            result += "\n";
            foreach (var i in items)
            {
                result += i.PrettyString(indent + 2);
            }

            return result;
        }

#if !OOXML
        public void pop(string format, object arg0)
#else
        public void pop([StringSyntax(StringSyntaxAttribute.CompositeFormat)] string format, object? arg0)
#endif
        {
            log(string.Format(format, arg0));
            pop();
        }

        public void pop(string str)
        {
            log(str);
            pop();
        }

        public void log(string v)
        {
            Debug.Assert(C.root != null);
            Debug.Assert(C.current_context != null);
            items.Add(new C(v));
        }

#if !OOXML
        public void log(string format, object arg0)
#else
        public void log([StringSyntax(StringSyntaxAttribute.CompositeFormat)] string format, object? arg0)
#endif
        {
            log(string.Format(format, arg0));
        }

#if !OOXML
        public void log(string format, object arg0, object arg1)
#else
        public void log([StringSyntax(StringSyntaxAttribute.CompositeFormat)] string format, object? arg0, object? arg1)
#endif
        {
            log(string.Format(format, arg0, arg1));
        }

        public void pop()
        {
            RdfTemplate.tw.WriteLine("done " + value);

            if (this == root)
            {
                root = null;
                C.current_context = null;
            }
            else
            {
                C p;

#if !OOXML
                p = (C)parent.Target;
                bool GotParent = p != null;
#else
                bool GotParent = parent.TryGetTarget(out p);
#endif
                Debug.Assert(GotParent);
                p.items.Remove(this);
                C.current_context = p;
            }
        }
    }
    

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

    public class UriLabelPair
    {
        public string uri;
        public string label;
        public UriLabelPair(string _uri, string _label)
        {
            uri = _uri;
            label = _label;
        }
    }

    public class UriLabelPairList : List<UriLabelPair>
    {
    }

    // a mapping from field to Pos
    public class FieldMap : Dictionary<INode, Pos>
    {
        public override string ToString()
        {
            return String.Join(", ", this.Select(kv => kv.Key.ToString()/*GetPropertyUri(kv.Key)*/ + " at " + kv.Value.ToString())); 
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

    public class RdfTemplateError : Exception
    {
        public RdfTemplateError(string message = "") : base(message)
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
        // the sheet currently being read or populated:
#if !OOXML
        private Worksheet _sheet;
        Excel.Application _app;
#else
        private IXLWorksheet _sheet;
        XLWorkbook _app;
#endif

        private INode _sheetsGroupTemplateUri;
        public string alerts;
        public static TextWriter tw;
        private readonly bool _isFreshSheet = true;
        // This is the main graph used throughout the lifetime of RdfTemplate. It is populated either with RdfTemplates.n3, or with response.n3. response.n3 contains also the templates, because they are sent with the request. We should maybe only send the data that user fills in, but this works:
        protected Graph _g;
        // here we put core request data that can be used to construct an example sheetset from a request:
        protected Graph _rg;

        // we generate some pseudo blank nodes, unique uris. But blank nodes work too.
        protected decimal _freeBnId = 0;


#if !OOXML
        public RdfTemplate(Excel.Application app)
        {
#if !DEBUG
            try
            {
#endif
            _app = app;
            Init();
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
#if !DEBUG
            try
            {
#endif
            _app = app;
            Init();
            _sheetsGroupTemplateUri = _g.CreateUriNode(sheetsTemplateQName);
#if !DEBUG
            }
            catch (Exception e)
            {
                MessageBox.Show("while initializing RdfTemplate(" + sheetsTemplateQName + "): " + e.Message, "LodgeIt", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw e;
            }
#endif
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
                MessageBox.Show("while initializing RdfTemplate(" + sheetsTemplateUri + "): " + e.Message, "LodgeIt", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw e;
            }
#endif
        }

#else

        public RdfTemplate(XLWorkbook app, string sheetsTemplateQName)
        {
            _app = app;
            Init();
            _sheetsGroupTemplateUri = u(sheetsTemplateQName);
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
#if !OOXML
            tw = System.Console.Out;
#endif
        }

#if !OOXML

        private void ErrMsg(string msg)
        {
            tw.WriteLine(msg);
            MessageBox.Show(msg, "LodgeiT", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
#else
        private void ErrMsg(string msg)
        {
            tw.WriteLine(msg);
            alerts += msg + "\n";
        }

        private void wl(string s)
        {
            tw.WriteLine(s);
        }
        
#endif

#if !OOXML

        public List<string> AvailableSheetSets(string rdf_templates)
        {
            C c = push("AvailableSheetSets");
            LoadTemplates(rdf_templates);
            List<string> result = new List<string>();
            foreach (var i in GetSubjects(u("rdf:type"), u("excel:sheet_set")))
                result.Add(i.AsValuedNode().AsString());
            c.pop();
            return result;
        }
        public UriLabelPairList ExampleSheetSets(string rdf_templates)
        {
            LoadTemplates(rdf_templates);
            UriLabelPairList result = new UriLabelPairList();
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
        
        
        private C push(string value)
        {
            tw.WriteLine(value + "...");
            C c = new C(value);
            C.current_context = c;
            return c;
        }
    
        private C push(string format, params object[] args)
        {
            return push(String.Format(format, args));
        }

        
        private bool CreateRdfEndpointRequestFromSheetGroupData()
        {
            C c = push("find worksheets relevant for {0} and generate structured data", _sheetsGroupTemplateUri);
            IEnumerable<INode> known_sheets = GetListItems(_sheetsGroupTemplateUri, "excel:sheets");
            var extracted_instances_by_sheet_type = new Dictionary<INode, IList<SheetInstanceData>>();
            if (!ExtractDataInstances(known_sheets, ref extracted_instances_by_sheet_type))
                return false;
            if (!make_sure_all_non_optional_sheets_are_present(known_sheets, extracted_instances_by_sheet_type))
                return false;
            if (!AssertRequest(extracted_instances_by_sheet_type))
                return false;
            c.pop();
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
            Assert(_g, u(":request"), u("l:client_version"), _g.CreateLiteralNode("3"));
            //Assert(_g, u(":request"), u("l:client_git_info"), _g.CreateLiteralNode(Properties.Resources.ResourceManager.GetObject("repo_status").ToString().Replace("\n", Environment.NewLine)));
            return true;
        }

        private bool ExtractDataInstances(IEnumerable<INode> known_sheets, ref Dictionary<INode, IList<SheetInstanceData>> extracted_instances_by_sheet_type)
        {
#if !OOXML
            foreach (Excel.Worksheet sheet in _app.Worksheets)
#else
            foreach (var sheet in _app.Worksheets)
#endif
            {
                    _sheet = sheet;
                if (!ScanSheet(known_sheets, extracted_instances_by_sheet_type)) return false;
            }
            return true;
        }

        private bool ScanSheet(IEnumerable<INode> known_sheets, Dictionary<INode, IList<SheetInstanceData>> extracted_instances_by_sheet_type)
        {
            C c = push("scan sheet '{0}'", _sheet.Name);
            if (GetCellValueAsString2(new Pos { col = 'A', row = 1 }).ToLower() != "sheet type:")
            {
                c.pop("sheet '{0}' does not have sheet type header", _sheet.Name);
                return true;
            }

            string sheet_type_uri_string = GetCellValueAsString2(new Pos { col = 'B', row = 1 });
            INode sheet_type_uri = _g.CreateUriNode(new Uri(sheet_type_uri_string));
            c.log("sheet '{0}' has sheet type: '{0}'", _sheet.Name, sheet_type_uri);
            
            if (!known_sheets.Contains(sheet_type_uri))
            {
                ErrMsg("unexpected sheet type: `" + sheet_type_uri_string + "`.");
                return false; //die
            }

            INode record_instance = null;
            INode sheet_template = GetObject(sheet_type_uri, u("excel:root"));
                
            if (!ExtractRecordByTemplate(sheet_template, ref record_instance))
                return false; //die
            
            Assert(_g, record_instance, u("excel:sheet_type"), sheet_type_uri);
            
            // todo as defaultdict: extracted_instances_by_sheet_type[sheet_type_uri].Add(new SheetInstanceData(_sheet.Name, record_instance));
            if (!extracted_instances_by_sheet_type.ContainsKey(sheet_type_uri))
                extracted_instances_by_sheet_type[sheet_type_uri] = new List<SheetInstanceData>();
            extracted_instances_by_sheet_type[sheet_type_uri].Add(new SheetInstanceData(_sheet.Name, record_instance));

            c.pop();
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
            try
            {
                LoadTemplates(UpdatedRdfTemplates);
                return this.CreateRdfEndpointRequestFromSheetGroupData();
            }
            catch (RdfTemplateError e)
            {
                FailReturn(e);
                return false;
            }
        }

        private void FailReturn(RdfTemplateError e)
        {
            PopulateAlertsFromTrace(e);
        }

        private void PopulateAlertsFromTrace(RdfTemplateError e)
        {
            alerts = "during:\n" + C.root.PrettyString() + "\nerror:\n" + e.Message;
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
#if !OOXML
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
#endif
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
#if !OOXML
        public void WriteFirstRow(INode sheet_decl)
        {
            WriteString(new Pos { col = 'A', row = 1 }, "sheet type:");
            //_sheet.Range["A1"].AddComment("blablabl\nablablabla");
            WriteString(new Pos { col = 'B', row = 1 }, sheet_decl.ToString());
        }
#endif
        public bool ExtractRecordByTemplate(INode template, ref INode individual)
        {
            C c = push("extract table at '{0}'", GetPos(template).ToString());

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
            c.pop();
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
            C c = push("ExtractRecord {0}", item_offset);
            var cls = GetClass(template);
            var values = new Dictionary<INode, RdfSheetEntry>();
            var cell_positions = GetRecordCellPositions(template, map, item_offset);
            if (!ReadCellValues(cell_positions, ref values))
            {
                c.pop();
                return false;
            }

            if (!ReadSubTemplates(template, ref values))
            {
                c.pop();
                return false;
            }

            if (!values.Any())
            {
                if (isRequired)
                {
                    String msg = "no values detected in template " + GetTitle(template) + ". Expected:\n";
                    foreach (KeyValuePair<INode, Pos> mapping in cell_positions)
                        msg += FieldTitles(mapping.Key).First() + " at " + mapping.Value.Cell + "\n";
                    ErrMsg(msg);
                    //should throw here?
                }
                c.pop("no values");
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
            c.pop();
            return true;
        }

        protected bool ReadCellValues(FieldMap cell_positions, ref Dictionary<INode, RdfSheetEntry> values)
        {
            C c = push("ReadCellValues {0}", cell_positions.ToString());

            foreach (KeyValuePair<INode, Pos> mapping in cell_positions)
            {
                INode field = mapping.Key;
                INode obj = null;
                Pos pos = cell_positions[field];
                IEnumerable<INode> types = GetObjects(field, "excel:type");

                if (!ReadCellAsType(pos, field, types, ref obj))
                {
                    c.pop();
                    return false; //either parsing error or possibly end of sheet
                }

                if (obj != null)
                    values[GetPropertyUri(field)] = new RdfSheetEntry(AssertValue(_rg, obj), pos);
                else if (!FieldIsOptional(field))
                {
                    ErrMsg("missing required field in " + _sheet.Name + " at " + pos.ToString() + ": " + FieldTitles(field).First());
                    c.pop();
                    return false;
                }
            }
            c.pop();
            return true;
        }

        private bool FieldIsOptional(INode field)
        {
            return BoolObjectWithDefault(field, u("excel:optional"), true);
        }

        INode DetermineTypeToReadCellAs(IEnumerable<INode> types)
        {
            if (!types.Any())
                return u("xsd:string");
            else
                return types.First();//todo shouldn't this use One?? Are there situations where there is more than one type specified but we want the first one always? Is this basically OneWithDefault?
        }
        bool ReadCellAsType(Pos pos, INode field, IEnumerable<INode> types, ref INode obj)
        {

            INode type = DetermineTypeToReadCellAs(types);
            
            /* xsd:decimal and integer are, here, regarded as being "optional" in the sense of a nullable value, an empty cell is not regarded as an error, but "represented" by a missing value.
             the types ("xsd:decimal") should eventually be changed to express that ("l:optional_decimal").
             But also a new type specifically for monetary values should be introduced, that deals in a special way with formatting. A decimal in the role of, say, percents, would not be entered with a $ sign.             
             see also "datatypes" in RdfTemplates.n3
             
             */
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
                if (!ReadOptionalDatetime(pos, ref obj))
                    return false;
            }
            else if (type.Equals(u("excel:uri")))
            {
                obj = GetUriObj(pos, field); //possibly throws
            }
            else if (type.Equals(u("xsd:string")))
            {
                string contents = null;
                CellReadingResult status = GetCellValueAsString(pos, ref contents);
                if (status == CellReadingResult.Ok)
                    obj = contents.ToLiteral(_g);
                else if (status == CellReadingResult.Error) // todo: maybe added for end of sheet data, but i'm not sure it makes sense / really happens
                    return false;
            }
            else
                throw new Exception("RDF template error: excel:type not recognized: " + type.ToString());
            return true;
        }

        public bool ReadOptionalInt(Pos pos, ref INode obj)
        /*
        return true if cell is empty
    set obj and return true on successful parse	
    show messagebox and return false on parse error.
        */
        /*this was plain wrong. If the cell contains a numeric value, and the column is not wide enough, ##### will be displayed, and then we'd be trying to parse that.
         also, why trim a $ off an integer? even if we wanted to have fields of type "monetary integer", it should be a different datatype.
         */ 
        {
            
#if !OOXML
            Excel.Range rng = _sheet.Range[pos.Cell];

            string txt =  rng.Text;
            txt = txt.Trim();
            if (txt == "")
                return true;

            try
            {
                int result = ((IConvertible)rng.Value2).ToInteger(null);
                obj = result.ToLiteral(_g);
                return true;
            }
            catch (System.FormatException e)
            {
                ErrMsg("error reading integer number in " + _sheet.Name + " at " + pos.Cell + ", got: \"" + txt + "\", error: " + e.Message);
                throw new RdfTemplateError();
            }

#else
            int result=123;//fixme
            var rng = _sheet.Cell(pos.Cell);
            
            string txt = GetCellValueAsString2(pos);
            if (txt == "")
                return true;

            try
            {
                /*not sure what the semantics of this attempted conversion are. Cells cannot technically contain integers, not sure how ClosedXML deals with this.
                 would this throw an error if the number was big enough not to fit in the range of integers that are exactly representable as doubles (or whatever actually excel uses to store cell values)?
                 do we need to possibly check that it's within that range?
                 we cannot just round whatever comes, because users could enter decimal number in an int field, and we don't want to silently round those.   
                 */
                result = (int)rng.Value;
            }
            catch (InvalidCastException e)
            {
                ErrMsg("error reading decimal in " + _sheet.Name + " at " + pos.Cell + ", got: \"" + txt + "\", error: " + e.Message);
                throw new RdfTemplateError();
            }

            obj = result.ToLiteral(_g);
            return true;
#endif
        }

    
    
    private bool ReadOptionalDecimal(Pos pos, ref INode obj)
        /*
        return true if cell is empty
        set obj and return true on successful parse	
        show messagebox and return false on parse error.
        */
        {
            
#if !OOXML
            Excel.Range rng = _sheet.Range[pos.Cell];

            string txt =  rng.Text;
            txt = txt.Trim();
            if (txt == "")
                return true;
            
            try
            {
                decimal result = ((IConvertible)rng.Value2).ToDecimal(null);
                obj = result.ToLiteral(_g);
                return true;
            }
            catch (System.FormatException e)
            {
                ErrMsg("error reading decimal in " + _sheet.Name + " at " + pos.Cell + ", got: \"" + txt + "\", error: " + e.Message);
                throw new RdfTemplateError();
            }

#else
            double result=123;//fixme
            var rng = _sheet.Cell(pos.Cell);

            string txt = GetCellValueAsString2(pos);
            if (txt == "")
                return true;
            
            try
            {
                rng.GetDouble();
            }
            catch (InvalidCastException e)
            {
                ErrMsg("error reading decimal in " + _sheet.Name + " at " + pos.Cell + ", got: \"" + txt +
                       "\", error: " + e.Message);
                throw new RdfTemplateError();
            }
            
            obj = result.ToLiteral(_g);
            return true;
#endif
        } 
    
    private bool ReadOptionalDatetime(Pos pos, ref INode obj)
        /*
        return true if cell is empty
        set obj and return true on successful parse	
        on parse error, return true and assert obj as a string.
        */
    {
#if !OOXML

??
            Excel.Range rng = _sheet.Range[pos.Cell];

            string txt =  rng.Text;
            txt = txt.Trim();
            if (txt == "")
                return true;

// throws what? we probably expect RdfTemplateErrror up the stack.        
DateTime contents = ExporttoXMLBase.GetCellAsDate(_sheet, pos.Cell);
        
        if (contents != DateTime.MinValue)
            obj = contents.Date.ToLiteral(_g);
        else
        {
            string contents_str = GetCellValueAsString2(pos); 
            if (contents_str != "") // if there was text but it couldn't be parsed, pass it on as string // do we take advantage of this anywhere on the prolog side?
                obj = contents_str.ToLiteral(_g);
        }
        return true;
#else
        
        var rng = _sheet.Cell(pos.Cell);

        string txt = GetCellValueAsString2(pos);
        if (txt == "")
            return true;

        DateTime result;

        try
        {
            result = rng.GetDateTime();
        }
        catch (InvalidCastException e)
        {
            obj = txt.ToLiteral(_g);
            return true;
        }

        obj = result.ToLiteral(_g);
        return true;
#endif
    }
    
    
        public string GetCellValueAsString2(Pos pos)
        {
            string value = "";
            GetCellValueAsString(pos, ref value);
            return value;
        }
        public CellReadingResult GetCellValueAsString(Pos pos, ref string result)
        {
#if !OOXML
            Range rng = _sheet.get_Range(pos.Cell);
            if (rng.Value2 != null)
                result = Convert.ToString(rng.Value2);
            else
                result = Convert.ToString(rng.Text);
            if (result == null)
                return CellReadingResult.Empty;
            result = result.Trim();
            if (result.Length == 0)
                return CellReadingResult.Empty;
            return CellReadingResult.Ok;
#else
            if (_sheet.Cell(pos.Cell).TryGetValue(out result))
            {
                result = result.Trim();
                if (result.Length == 0)
                    return CellReadingResult.Empty;
                return CellReadingResult.Ok;
            }
            else
            {
                ErrMsg("error in " + _sheet.Name + " " + pos.Cell);
                return CellReadingResult.Error;
            }
#endif
        }

        public void Assert(IGraph g, INode s, INode p, INode o)
        {
            /*dotnetrdf 3.0 breaking(?) change, nodes are no longer specific to individual graphs. Can we upgrade to 3.0? */
            g.Assert(new Triple(Tools.CopyNode(s, g), Tools.CopyNode(p, g), Tools.CopyNode(o, g)));
            //g.Assert(new Triple(s, p, o));
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

            string contents = GetCellValueAsString2(pos);
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
                string sKey = GetCellValueAsString2(pos2).ToLower();
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

#if !OOXML
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
            TemplateGenerator.AddBoldValueBorder(_sheet, pos.Cell, title);
            pos.row++;
            foreach (var field in GetFields(template))
            {
                if (!GetObjects(field, "excel:template").Any())
                {
                    string cell_title = FieldTitles(field).First();
                    TemplateGenerator.AddBoldValueBorder(_sheet, pos.Cell, cell_title);
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
#endif
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
#if !OOXML
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
            C c = push("LoadTemplates");

#if !DEBUG
            try
            {
#endif
            StreamReader reader;
            if (UpdatedRdfTemplates != null && UpdatedRdfTemplates != "")
                /*these live-updated templates is probably the only source we should support */
                reader = new StreamReader(new MemoryStream(Encoding.UTF8.GetBytes(UpdatedRdfTemplates)));
            else
#if JINDRICH_DEBUG
                reader = new StreamReader(File.OpenRead(@"C:\Users\kokok\source\repos\LodgeITSmart\LodgeiTSmart\LodgeiTSmart\Resources\RdfTemplates.n3"));
#else
#if !OOXML
                reader = new StreamReader(new MemoryStream((byte[])Properties.Resources.ResourceManager.GetObject("RdfTemplates")));
#else
                reader = new StreamReader(File.OpenRead(/*Environment.GetEnvironmentVariable("CSHARPSERVICES_DATADIR") + "/" +  */ "RdfTemplates.n3"));
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
            c.pop();
        }
#if !OOXML
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
