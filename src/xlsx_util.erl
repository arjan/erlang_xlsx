-module(xlsx_util).

-export([new/0, write/2, mkdir/2, get_output_stream/2, write/3, get_sheets/1]).

-record(xlsx, {tmp, files=[], sheets=[]}).

%% @doc Create new xlsx doc
new() ->
    Dir = tempfile(),
    ok = file:make_dir(Dir),
    #xlsx{tmp=Dir}.

get_sheets(#xlsx{sheets=Sheets}) ->
    Sheets.

%% @doc Write zip file contents
write(#xlsx{tmp=Tmp, files=Files}, OutFile) ->
    {ok, CurrentDir} = file:get_cwd(),
    file:set_cwd(Tmp),
    {ok, Path} = zip:create(filename:join(CurrentDir, OutFile), Files),
    os:cmd("rm -rf " ++ Tmp),
    file:set_cwd(CurrentDir),
    {ok, Path}.

mkdir(#xlsx{tmp=Tmp}, Dir) ->
    lists:foldl(
      fun(Part, Path) ->
              D = filename:join(Path, Part),
              case filelib:is_dir(D) of
                  true ->
                      nop;
                  false ->
                      file:make_dir(D)
              end,
              D
      end,
      Tmp,
      string:tokens(Dir, "/")).

get_output_stream(X=#xlsx{tmp=Tmp, files=Files}, RelPath) ->
    mkdir(X, filename:dirname(RelPath)),
    Path = filename:join(Tmp, RelPath),
    {ok, F} = file:open(Path, [write]),
    {ok, {F, X#xlsx{files=[RelPath|Files]}}}.

write(X, RelPath, Bytes) ->
    {ok, {F, X2}} = get_output_stream(X, RelPath),
    ok = file:write(F, Bytes),
    ok = file:close(F),
    {ok, X2}.

%% helpers

tempfile() ->
    {A,B,C}=erlang:now(),
    filename:join(temppath(), lists:flatten(io_lib:format("xlsx-~s-~p.~p.~p",[node(),A,B,C]))).


%% @doc Returns the path where to store temporary files.
temppath() ->
    lists:foldl(fun(false, Fallback) -> Fallback;
                   (Good, _) -> Good end,
                "/tmp",
                [os:getenv("TMP"), os:getenv("TEMP")]).
              
