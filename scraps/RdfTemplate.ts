'use strict';


	// a mapping from field to Pos
	class FieldMap{};// Dictionary<INode, Pos>;

	/// abstraction of excel cell coordinates
	enum CellReadingResult
	{
		Ok,
		Error,
		Empty
	}

	class Pos
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

	// contains the value obtained at a cell, and the cell's position
	class RdfSheetEntry
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

	/// <summary>
	/// schema-directed UI
	/// in future, we should replace the data structure a bit, probably use http://datashapes.org/forms.html
	/// 
	/// </summary>
	/// note on control flow style used:
	/// reporting errors to user is the competence of the function that detects the error. Error is displayed and false is returned.
	/// object lifecycle:
	/// construct, call one of: GenerateTemplate, ExtractSheetGroupData or DisplayResponse, and dispose
	class RdfTemplate : TemplateGenerator
	{
		
		
		private string _sheetsGroupTemplateUri;
		
		
		// the sheet currently being read or populated
		private Worksheet _sheet;
		Excel.Application _workbook;
		private readonly bool _isFreshSheet = true;
		/*
            Currently there is only one graph throughout the lifetime of RdfTemplate. It is populated either with RdfTemplates.n3, or with response.n3. response.n3 contains the templates, because they are sent with the request.
            We should maybe only send the data that user fills in, but this works.
        */
		protected Graph _g;
		protected decimal _freeBnId = 0;

		public RdfTemplate(Excel.Application workbook, string sheetsTemplateUri)
		{
			_sheetsGroupTemplateUri = sheetsTemplateUri;
			_workbook = workbook;
		}

		public void CreateSheetsFromTemplate(string rdf_templates)
		{
			LoadTemplates(rdf_templates);

			foreach example_sheet_info in example_sheet_sets[_sheetsGroupTemplateUri].example_has_sheets
			{
				const sheet_decl = sheet_types[example_sheet_info.has_sheet]
				string sheet_name = GetSheetNamePrefix(sheet_decl);
				var template = sheet_decl.root
				var doc = example_sheet_info.example_doc
				_sheet = NewWorksheet(sheet_name, GetMultipleSheetsAllowed(sheet_decl));
				if (_sheet != null)
				{
					WriteFirstRow(sheet_decl);
					WriteData(template, doc);
					_sheet.Columns.AutoFit();
				}
			}
		}

		public bool ExtractSheetGroupData(string UpdatedRdfTemplates)
		{
			LoadTemplates(UpdatedRdfTemplates);
			var known_sheets = sheet_sets[_sheetsGroupTemplateUri].sheets
			var extracted_instances_by_sheet_type = {}
			foreach (Excel.Worksheet sheet in _workbook.Worksheets)
			{
				_sheet = sheet;

				// ignore any sheet that does not have the type header
				if (GetCellValue({ col:'A', row:1 }).Trim().ToLower() != "sheet type:")
					continue;

				string sheet_type_uri = GetCellValue(new Pos { col = 'B', row = 1 }).Trim();
				
				if (!known_sheets.Contains(sheet_type_uri))
				{
					MessageBox.Show("unknown sheet type: " + sheet_type_uri + ", ignoring.", "LodgeIt");
					return false;
				}
				var sheet_type = sheet_types[sheet_type_uri]
				INode instance = null;
				INode sheet_template = sheet_type.root
				if (!ExtractSheetData(sheet_template, ref instance))
					return false;
				/*let's worry about passing the data back later
				_g.Assert(instance, u("excel:has_sheet_name"), AssertValue(_g.CreateLiteralNode(sheet.Name)));
				_g.Assert(instance, u("excel:sheet_type"), sheet_type_uri);
				if (!extracted_instances_by_sheet_type.ContainsKey(sheet_type_uri))
					extracted_instances_by_sheet_type[sheet_type_uri] = new List<INode>();
				extracted_instances_by_sheet_type[sheet_type_uri].Add(instance);*/
			}
			/* make sure all non-optional sheets are present */
			foreach (INode known_sheet in known_sheets)
			{
				known_sheet = sheet_types[known_sheet]
				if (known_sheet.optional)
					continue;
				if (extracted_instances_by_sheet_type.ContainsKey(known_sheet))
					continue;
				string msg = "sheet \"" + GetSheetNamePrefix(known_sheet) + "\" (" + known_sheet.ToString() + ") not found.";
				MessageBox.Show(msg, "LodgeIt");
				return false;
			}
			foreach (KeyValuePair<INode, IList<INode>> kv in extracted_instances_by_sheet_type)
			{
				INode sheet_template = GetObject(kv.Key, u("excel:root"));
				INode result;
				if (!GetMultipleSheetsAllowed(kv.Key))
					result = kv.Value.First();
				else
					result = _g.AssertList(kv.Value);

				/* here we use sheet_template as a predicate, this is somewhat questionable. 
				For example, :request smsf_ui:members_sheet something.
				As with many other predicates in RDF and logic, a predicate that doesn't consist of a verb us to be understood as implying a verb, usually "is" or "has". In this case, it's "has". :request has smsf_ui:members_sheet something */
				_g.Assert(new Triple(u(":request"), sheet_template, result));


			}
			_g.Assert(new Triple(u(":request"), u("l:client_version"), _g.CreateLiteralNode("2")));
			return true;
		}

		private bool GetMultipleSheetsAllowed(INode sheet_decl)
		{
			return sheet_decl.multiple_sheets_allowed || false
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
		public void DisplayData(StreamReader data, bool inPlace)
		{
			LoadRdf(data);
			foreach (var triple in _g.GetTriplesWithPredicate(u("excel:has_sheet_name")))
			{
				var doc = triple.Subject;
				string sheet_name = GetValue(triple.Object).AsValuedNode().AsString();
				var template = GetObject(doc, "excel:template");
				if (inPlace)
					_sheet = SheetByName(sheet_name);
				else
					_sheet = NewWorksheet(sheet_name, true);
				WriteFirstRow(GetObject(doc, u("excel:sheet_type")));
				WriteData(template, doc);
				_sheet.Columns.AutoFit();
			}
		}

		public void WriteFirstRow(INode sheet_decl)
		{
			WriteString(new Pos { col = 'A', row = 1 }, "sheet type:");
			WriteString(new Pos { col = 'B', row = 1 }, sheet_decl.ToString());
		}

		public bool ExtractSheetData(INode template, ref INode individual)
		{
			var map = new Dictionary<INode, Pos>();
			if (!HeadersMapping(template, ref map))
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
						individuals.Add(item);
						numEmptyRows = 0;
					}
					else
						numEmptyRows++;
					item_offset++;
				} while (numEmptyRows < 5);
				var rdf_list = _g.AssertList(individuals);
				individual = Bn("list");
				_g.Assert(individual, u("rdf:value"), rdf_list);
			}
			_g.Assert(individual, u("excel:template"), template);
			return true;
		}

		protected bool ExtractRecord(INode template, FieldMap map, int item_offset, ref INode record, bool isRequired)
		{
			var cls = GetClass(template);
			bool is_horiz = GetIsHorizontal(template);
			/* maybe just return null record on failure */
			record = Bn("record");
			var values = new Dictionary<INode, RdfSheetEntry>();
			var cell_positions = new Dictionary<INode, Pos>();
			foreach (KeyValuePair<INode, Pos> mapping in map)
			{
				Pos pos = mapping.Value.Clone();
				if (is_horiz)
					pos.col += item_offset;
				else
					pos.row += item_offset;
				cell_positions[mapping.Key] = pos;
			}
			AssertPosFlat(record, cell_positions.First().Value);
			_g.Assert(new Triple(record, u("excel:sheet_name"), _sheet.Name.ToLiteral(_g)));
			foreach (var field in GetFields(template))
			{
				var field_template = GetObjects(field, "excel:template");
				INode obj = null;
				if (field_template.Any())
				{
					if (!ExtractSheetData(One(field_template), ref obj))
						return false;
					if (obj != null)
						values[GetPropertyUri(field)] = new RdfSheetEntry(obj, null);

				}
				else if (cell_positions.ContainsKey(field))
				{
					var pos = cell_positions[field];
					var types = GetObjects(field, "excel:type");
					INode type;
					if (!types.Any())
						type = u("xsd:string");
					else
						type = types.First();
					if (type.Equals(u("xsd:decimal")))
					{
						decimal? contents = ExporttoXMLBase.GetCellValueAsDecimalNullable(_sheet, pos.Cell);
						if (contents != null)
						{
							obj = contents.Value.ToLiteral(_g);
						}
					}
					else if (type.Equals(u("xsd:integer")))
					{
						decimal? contents = ExporttoXMLBase.GetCellValueAsIntegerNullable(_sheet, pos.Cell);
						if (contents != null)
						{
							obj = contents.Value.ToLiteral(_g);
						}
					}
					else if (type.Equals(u("xsd:dateTime")))
					{
						DateTime contents = ExporttoXMLBase.GetCellAsDate(_sheet, pos.Cell);
						if (contents != DateTime.MinValue)
							obj = contents.Date.ToLiteral(_g);
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
						string contents=null;
						CellReadingResult status = GetCellValueAsString(pos, ref contents);
						if (status == CellReadingResult.Ok)
							obj = contents.ToLiteral(_g);
						else if (status == CellReadingResult.Error)
							return false;
					}
					else
						throw new Exception("excel:type not recognized: " + type.ToString());
					if (obj != null)
						values[GetPropertyUri(field)] = new RdfSheetEntry(AssertValue(obj), pos);
					else if (!BoolObjectWithDefault(field, u("excel:optional"), true))
					{
						MessageBox.Show("missing required field in " + _sheet.Name + " at " + pos.Cell);
						return false;
					}
				}
			}
			if (!values.Any())
			{
				if (isRequired)
				{
					String msg = "no values detected in section " + GetTitle(template) + ". Expected:\n";
					foreach (KeyValuePair<INode, Pos> mapping in cell_positions)
						msg += FieldTitles(mapping.Key).First() + " at " + mapping.Value.Cell + "\n";
					MessageBox.Show(msg, "Lodge iT");
				}
				return false;
			}
			_g.Assert(new Triple(record, u("rdf:type"), cls));
			foreach (KeyValuePair<INode, RdfSheetEntry> entry in values)
			{
				_g.Assert(new Triple(record, entry.Key, entry.Value._obj));
				if (entry.Value._pos != null)
					AssertPosFlat(entry.Value._obj, entry.Value._pos);
				_g.Assert(new Triple(entry.Value._obj, u("excel:sheet_name"), _sheet.Name.ToLiteral(_g)));
			}
			return true;
		}
		/*
		todo
		public CellReadingResult GetCellValueAsInteger(Pos pos, ref int result)
		{
			Range rng = _sheet.get_Range(pos.Cell, pos.Cell);
			if (rng.Value2 is Int32)
			{
				MessageBox.Show("error in " + _sheet.Name + " " + pos.Cell);
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
		public CellReadingResult GetCellValueAsString(Pos pos, ref string result)
		{
			Range rng = _sheet.get_Range(pos.Cell, pos.Cell);
			if (rng.Value2 is Int32)
			{
				MessageBox.Show("error in " + _sheet.Name + " " + pos.Cell);
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
			result = sValue.Trim('$');
			return CellReadingResult.Ok;
		}

		public INode AssertValue(INode obj)
		{
			var value = Bn("value");
			//_g.Assert(new Triple(value, u("rdf:type"), u("l:value")));
			_g.Assert(new Triple(value, u("rdf:value"), obj));
			//_g.Assert(new Triple(value, _g.CreateUriNode("l:unit"), _g.CreateVariableNode("u")));
			return value;
		}
		public INode GetValue(INode s)
		{
			return GetObject(s, u("rdf:value"));
		}
		private INode GetUriObj(Pos pos, INode field)
		{
			INode obj;
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
				MessageBox.Show(msg, "LodgeIT");
				throw;
			}
		}

		/*find a header of each field, produce a mapping from field to Pos*/
			protected bool HeadersMapping(INode template, ref FieldMap result)
		{
			var pos = GetPos(template);
			bool is_horiz = GetIsHorizontal(template);
			// skip the section title
			pos.row += 1; 
			int nested_templates_count = 0;
			var fields = GetFields(template);
			Trace.Assert(fields.Count() > 0);
			List<string> expected_header_items = new List<string>();
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
				foreach (int item in Enumerable.Range(0, 100))
				{
					Pos pos2 = pos.Clone();
					if (is_horiz)
						pos2.row += item;
					else
					{
						pos2.col += item;
						if (item > 25)
							break;
					}		
					string sKey = GetCellValue(pos2).ToLower();
					if (titles.Contains(sKey, StringComparer.OrdinalIgnoreCase))
					{
						result[field] = pos2;
						break;
					}
				}
			}
			if (nested_templates_count == 0 && !result.Any())
			{
				string msg = "No headers found, for template \"" + GetTitle(template) + "\", sheet \"" + _sheet.Name + "\".\nexpected: ";
				foreach (string i in expected_header_items)
					msg += i + ", ";
				msg += "starting from " + pos.Cell;
				MessageBox.Show(msg, "LodgeIT");
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
			}
			else
				// it's an URI
				WriteString(pos, GetLabels(value_value).First());
		}

		void WriteDate(Pos pos, DateTime dt)
		{
			var rng = _sheet.Range[pos.Cell];
			rng.NumberFormat = "dd/mm/yyyy";
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
			_g.Assert(new Triple(subject, u("excel:col"), pos.GetExcelColumnName().ToLiteral(_g)));
			_g.Assert(new Triple(subject, u("excel:row"), pos.row.ToString().ToLiteral(_g)));
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
			foreach (Triple triple in _g.GetTriplesWithSubjectPredicate(subject, predicate))
				yield return triple.Object;
		}
		protected IEnumerable<INode> GetObjects(INode subject, string predicate)
		{
			return GetObjects(subject, u(predicate));
		}
		protected INode u(string uri)
		{
			return _g.CreateUriNode(uri);
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
		protected INode Bn(string id_base)
		{
			return u(":bn_" + id_base + (_freeBnId++).ToString());
		}


		public static string SerializeGraph(IGraph g)
		{
/*#if DEBUG
			SaveAsRdfXml(g);
#endif*/
			Notation3Writer w = new Notation3Writer();
			return VDS.RDF.Writing.StringWriter.Write(g, w);
		}

#if DEBUG
		protected static void SaveAsRdfXml(IGraph g)
		{
			/*RdfXmlWriter rdfXmlWriter = new RdfXmlWriter();
			rdfXmlWriter.Save(g, "j:\\Example.rdf.xml");*/
		}
#endif
		// isMulti = true:
		//	create, or create with a different name, if a sheet with name sheet_name already exists
		// isMulti = false:	
		//  create sheet with name sheet_name, or fail
		private Worksheet NewWorksheet(string sheet_name, bool isMulti)
		{
			if (!isMulti && SheetByName(sheet_name) != null)
			{
				MessageBox.Show("sheet with that name already exists: " + sheet_name, "Lodge iT");
				return null;
			}
			Worksheet worksheet = _workbook.Sheets.Add();
			worksheet.Name = GetUniqueName(_workbook.Sheets, sheet_name);
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


		private bool GetIsHorizontal(INode template)
		{
			if (!IsMulti(template))
				return true;
			return BoolObjectWithDefault(template, u("excel:is_horizontal"), false);
		}
		public string Serialize()
		{
			return SerializeGraph(_g);
		}
		protected void LoadRdf(StreamReader data)
		{
			var parser = new Notation3Parser();
			parser.Load(_g, data);
		}
		protected void LoadTemplates(string UpdatedRdfTemplates)
		{
			StreamReader reader;
			if (UpdatedRdfTemplates != null)
				reader = new StreamReader(new MemoryStream(Encoding.UTF8.GetBytes(UpdatedRdfTemplates)));
			else
#if DEV
				reader = new StreamReader(File.OpenRead(@"C:\Users\kokok\source\repos\LodgeITSmart\LodgeiTSmart\LodgeiTSmart\Resources\RdfTemplates.n3"));
#else

				reader = new StreamReader(new MemoryStream((byte[])Properties.Resources.ResourceManager.GetObject("RdfTemplates")));
#endif
			LoadRdf(reader);
		}

		protected Worksheet SheetByName(string name)
		{
			foreach (Excel.Worksheet sheet in _workbook.Worksheets)
				if (name.Trim().ToLower() == sheet.Name.Trim().ToLower())
					return sheet;
			return null;
		}
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
