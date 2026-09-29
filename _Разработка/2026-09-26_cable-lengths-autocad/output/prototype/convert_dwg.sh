#!/bin/sh
# Как в чате читался DWG без AutoCAD: сборка LibreDWG из исходников и перевод DWG (AC1032 / AutoCAD 2018+) в DXF,
# дальше DXF читался библиотекой ezdxf (из исходников GitHub, т.к. pip-зеркало было недоступно).
git clone --depth 1 --recurse-submodules --shallow-submodules https://github.com/LibreDWG/libredwg
git clone --depth 1 https://github.com/mozman/ezdxf
cd libredwg
sh autogen.sh
./configure --disable-bindings --disable-docs --disable-werror
make -j2 -C src && make -j2 -C programs
cd ..
cp "03_Чертежи.dwg" plan.dwg
./libredwg/programs/dwg2dxf -y -o plan.dxf plan.dwg
# далее: PYTHONPATH=ezdxf/src:. python3 -c 'exec(open("extract.py").read()) ...'
