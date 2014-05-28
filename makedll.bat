
set RUBYINFO=-I%MRUBY%\include %MRUBY%\build\i386-mingw32\lib\libmruby.a -lws2_32

dlltool -e vbonmruby.exp -d vbonmruby.def

gcc -Wall -shared -o vbonmruby.dll vbonmruby.exp  vbonmruby.c %RUBYINFO% -Wl,--enable-stdcall-fixup
copy /Y vbonmruby.dll vbdemo\
