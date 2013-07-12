-module(xlsx_sheet).

-export([encode_sheet/1]).

encode_sheet(Rows) ->
    [
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
     "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
     "<sheetData>",
     [encode_row(Idx, Row)  || {Idx, Row} <- lists:zip(lists:seq(1, length(Rows)), Rows)],
     "</sheetData></worksheet>"
    ].

encode_row(Idx, Row) ->
    I = integer_to_list(Idx),
    ["<row r=\"", I, "\">",
     [begin
          {Kind, Content, Style} = encode(Cell),
          ["<c r=\"", column(Col), I, "\" t=\"", atom_to_list(Kind), "\"",
           " s=\"", integer_to_list(Style), "\">", Content, "</c>"
          ]
      end || {Col, Cell} <- lists:zip(lists:seq(0, length(Row)-1), Row)],
     "</row>"].
    
%% Given 0-based column number return the Excel column name as list.
column(N) ->
    column(N, []).
column(N, Acc) when N < 26 ->
    [(N)+$A | Acc];
column(N, Acc) ->
    column(N div 26-1, [(N rem 26)+$A|Acc]).

%% @doc Encode an Erlang term as an Excel cell.
encode(true) ->
    {b, "<v>1</v>", 6};
encode(false) ->
    {b, "<v>0</v>", 6};
encode(I) when is_integer(I) ->
    {n, ["<v>", integer_to_list(I), "</v>"], 3};
encode(F) when is_float(F) ->
    {n, ["<v>", float_to_list(F), "</v>"], 4};
% encode(Dt = {{_,_,_},{_,_,_}}) ->
%     {n, ["<v></v>"], 1};
encode(Str) ->
    {inlineStr, ["<is><t>", z_html:escape(z_convert:to_list(Str)), "</t></is>"], 5}.


% ms_epoch() ->
%     calendar:datetime_to_gregorian_seconds({{1904, 1, 1}, {0, 0, 0}}).
    
