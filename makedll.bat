
set RUBYINFO=-I%MRUBY%\include %MRUBY%\build\host\lib\libmruby.a -lws2_32

dlltool -e vbonmruby.exp -d vbonmruby.def

gcc -Wall -shared -o vbonmruby.dll vbonmruby.exp  vbonmruby.c %RUBYINFO%
copy /Y vbonmruby.dll vbdemo\
