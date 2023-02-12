```
$ sudo apt install build-essential cmake cxxtest python-pycryptopp libcrypto++-dev zlib1g-dev
```

```
$ cd repo
$ git clone https://github.com/gonrin/jsoncpp.git
$ git clone https://github.com/gonrin/pybind11.git
$ git clone https://github.com/gonrin/xlnt.git
```

## Build jsoncpp

```
$ cd jsoncpp
$ mkdir -p build/debug
$ cd build/debug
$ cmake -DCMAKE_BUILD_TYPE=debug -DBUILD_STATIC_LIBS=ON -DBUILD_SHARED_LIBS=ON -DARCHIVE_INSTALL_DIR=. -G "Unix Makefiles" ../..
$ sudo make install
$ make
```

## Build xlnt

```
$ sudo apt install zlibc
$ cd ../../../
$ cd xlnt
$ cmake .
$ sudo make install
```

## Build pybuild11

```
$ sudo apt install python3-dev python3.7-dev
$ sudo pip3 install pytest
$ cd ../pybind11
$ mkdir build
$ cd build
$ cmake ..
```

## Build xlrender

```
$ cd ../../xlrender

#python3.7
$ g++ -O3 -shared -std=c++14 -fpic -I../jsoncpp/include/ -I../pybind11/include/ -I./include/ -I/usr/include/ -I /usr/local/include/ -o xlrender.so xlrender.cpp `python3.7-config --cflags --ldflags` -L/usr/lib/crypto++ -lcrypto++ -ljsoncpp -lz  -lxlnt

#python3.8
$ g++ -O3 -shared -std=c++14 -fpic -I../jsoncpp/include/ -I../pybind11/include/ -I./include/ -I/usr/include/ -I /usr/local/include/ -o xlrender.so xlrender.cpp `python3.8-config --cflags --ldflags` -L/usr/lib/crypto++ -lcrypto++ -ljsoncpp -lz  -lxlnt
```
## shared libraries in /usr/local/lib


```
$ sudo ldconfig /usr/local/lib
```
