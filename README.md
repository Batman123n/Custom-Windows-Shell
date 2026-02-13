# This is my replacement for explorer.exe (windows shell).
To compile this monster you need to have json.hpp in same dir as shell02.cpp
# Here is compile command:
g++ -O3 -flto -DNDEBUG -municode -Wno-stringop-overread -o shell02.exe shell02.cpp -lgdi32 -lshell32 -lcomctl32 -lole32 -luuid -lgdiplus -ldwmapi -lpsapi -lPowrProf -mwindows -std=c++17 -s
