-module(xlsx).

-export([create/2]).

-record(xlsx, {tmp, files=[], sheets=[]}).

create(Sheets, OutFile) ->
    Xlsx0 = xlsx_util:new(),
    Xlsx1 = Xlsx0#xlsx{sheets=Sheets},
    {ok, Xlsx2} = lists:foldl(fun({Idx, {_Name, Rows, ColInfo}}, {ok, Acc}) ->
                                I = integer_to_list(Idx),
                                xlsx_util:write(
                                  Acc, "xl/worksheets/sheet" ++ I ++ ".xml",
                                  xlsx_sheet:encode_sheet(Rows, ColInfo))
                        end,
                        {ok, Xlsx1},
                              numbered_sheets(Xlsx1)),
    {ok, XlsxNew} = lists:foldl(
                      fun(F, {ok, X}) -> F(X) end,
                      {ok, Xlsx2},
                      [fun add_doc_props/1,
                       fun add_relationship_part/1,
                       fun add_styles/1,
                       fun add_workbook_relationship_part/1,
                       fun add_content_types/1,
                       fun add_workbook_part/1
                      ]),
    xlsx_util:write(XlsxNew, OutFile).


numbered_sheets(#xlsx{sheets=Sheets}) ->
    lists:zip(lists:seq(1,length(Sheets)), Sheets).

add_workbook_part(Xlsx) ->
    xlsx_util:write(
      Xlsx, "xl/workbook.xml",
      [<<"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
<workbookPr date1904=\"0\" />
<sheets>">>,
       lists:map(fun({Idx, {SheetName, _Rows, _ColInfo}}) ->
                         I = integer_to_list(Idx),
                         ["<sheet name=\"", z_html:escape(SheetName), "\" sheetId=\"", I, "\" r:id=\"sheet", I, "\"/>"]
                 end,
                 numbered_sheets(Xlsx)),
       "</sheets></workbook>"]).

add_workbook_relationship_part(Xlsx) ->
    xlsx_util:write(
      Xlsx, "xl/_rels/workbook.xml.rels",
      [
       "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
       "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
       "<Relationship Id=\"rId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
       lists:map(fun({Idx, _Sheet}) ->
                         I = integer_to_list(Idx),
                         ["<Relationship Id=\"sheet", I, "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet", I, ".xml\"/>"]
                 end,
                 numbered_sheets(Xlsx)),
       "<Relationship Id=\"rId99\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"xl/sharedStrings.xml\"/>"
      ]).

add_content_types(Xlsx) ->
    xlsx_util:write(
      Xlsx, "[Content_Types].xml",
      ["<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
       "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
       "<Override PartName=\"/_rels/.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
       "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>",
       "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>",
       "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>",
       "<Override PartName=\"/xl/_rels/workbook.xml.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
       "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>",

       lists:map(fun({Idx, _Sheet}) ->
                         I = integer_to_list(Idx),
                         ["<Override PartName=\"/xl/worksheets/sheet", I, ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"]
                 end,
                 numbered_sheets(Xlsx)),
       "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>",
       "</Types>"
      ]).

add_relationship_part(Xlsx) ->
    xlsx_util:write(
      Xlsx, "_rels/.rels",
      [
       "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
       "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
       "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
       "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
       "  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
       "</Relationships>"
      ]).

add_doc_props(Xlsx) ->
    {ok, X2} = xlsx_util:write(
      Xlsx, "docProps/core.xml",
      [
       "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
       "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
       "   <dcterms:created xsi:type=\"dcterms:W3CDTF\">2010-07-20T14:30:58.00Z</dcterms:created>"
       "   <cp:revision>0</cp:revision>"
       "</cp:coreProperties>"
      ]),
    xlsx_util:write(
      X2, "docProps/app.xml",
      ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
       "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
       "  <TotalTime>0</TotalTime>"
       "</Properties>"
     ]).

add_styles(Xlsx) ->
    xlsx_util:write(
      Xlsx, "xl/styles.xml",
      [
       "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
       "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
       "<numFmts count=\"7\">",
       "<numFmt formatCode=\"GENERAL\" numFmtId=\"164\"/>",
       "<numFmt formatCode=\"&quot;yes&quot;;&quot;yes&quot;;&quot;no&quot;\" numFmtId=\"170\"/>",
       "</numFmts>",
       "<fonts count=\"5\">",
       "  <font><name val=\"Arial\"/><family val=\"2\"/><sz val=\"10\"/></font>"
       "  <font><name val=\"Arial\"/><family val=\"0\"/><sz val=\"10\"/><b val=\"true\"/></font>"
       "  <font><name val=\"Arial\"/><family val=\"0\"/><sz val=\"10\"/><i val=\"true\"/></font>"
       "  <font><name val=\"Arial\"/><family val=\"0\"/><sz val=\"10\"/></font>"
       "  <font><name val=\"Arial\"/><family val=\"2\"/><sz val=\"10\"/></font>"
       "</fonts>"
       "<fills count=\"2\">"
       "  <fill><patternFill patternType=\"none\"/></fill>"
       "  <fill><patternFill patternType=\"gray125\"/></fill>"
       "</fills>"
       "<borders count=\"1\">"
       "  <border diagonalDown=\"false\" diagonalUp=\"false\"><left/><right/><top/><bottom/><diagonal/></border>"
       "</borders>"
       "<cellStyleXfs count=\"20\">"
       "  <xf applyAlignment=\"true\" applyBorder=\"true\" applyFont=\"true\" applyProtection=\"true\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"164\">"
       "    <alignment horizontal=\"general\" indent=\"0\" shrinkToFit=\"false\" textRotation=\"0\" vertical=\"bottom\" wrapText=\"false\"/>"
       "    <protection hidden=\"false\" locked=\"true\"/>"
       "  </xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"43\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"41\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"44\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"42\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"9\"></xf>"
       "  </cellStyleXfs>"
       "<cellXfs count=\"10\">"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"4\" numFmtId=\"164\" xfId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"4\" numFmtId=\"22\" xfId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"4\" numFmtId=\"15\" xfId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"4\" numFmtId=\"1\" xfId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"4\" numFmtId=\"2\" xfId=\"0\"></xf>"
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"4\" numFmtId=\"49\" xfId=\"0\"></xf>"
       %% boolean
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"false\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"4\" numFmtId=\"170\" xfId=\"0\"></xf>"
       %% string bold
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"1\" numFmtId=\"49\" xfId=\"0\"></xf>"
       %% string italic
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"0\" fontId=\"2\" numFmtId=\"49\" xfId=\"0\"></xf>"

       %% string filled
       "  <xf applyAlignment=\"false\" applyBorder=\"false\" applyFont=\"true\" applyProtection=\"false\" borderId=\"0\" fillId=\"1\" fontId=\"0\" numFmtId=\"49\" xfId=\"0\"></xf>"
       
       "</cellXfs>"
       "<cellStyles count=\"6\"><cellStyle builtinId=\"0\" customBuiltin=\"false\" name=\"Normal\" xfId=\"0\"/>"
       "  <cellStyle builtinId=\"3\" customBuiltin=\"false\" name=\"Comma\" xfId=\"15\"/>"
       "  <cellStyle builtinId=\"6\" customBuiltin=\"false\" name=\"Comma [0]\" xfId=\"16\"/>"
       "  <cellStyle builtinId=\"4\" customBuiltin=\"false\" name=\"Currency\" xfId=\"17\"/>"
       "  <cellStyle builtinId=\"7\" customBuiltin=\"false\" name=\"Currency [0]\" xfId=\"18\"/>"
       "  <cellStyle builtinId=\"5\" customBuiltin=\"false\" name=\"Percent\" xfId=\"19\"/>"
       "</cellStyles>"
       "</styleSheet>"
      ]).       
