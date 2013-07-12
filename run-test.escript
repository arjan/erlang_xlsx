#!/usr/bin/env escript

main([]) ->
    true = code:add_pathz(filename:dirname(escript:script_name()) 
                          ++ "/ebin"),
    true = code:add_pathz(filename:dirname(escript:script_name()) 
                          ++ "/deps/z_stdlib/ebin"),

    xlsx:create([{"Sheet nr 1",
                  [[1,2,3,true], ["hallo"]]},
                 {"Sheet 2",
                  [["foo", 1.23]]}],
                "test.xlsx").

