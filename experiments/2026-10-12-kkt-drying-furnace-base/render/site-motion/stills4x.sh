#!/bin/bash
# 장면 시각 t → 프레임 (t+3)*15
cd "$(dirname "$0")"; mkdir -p stills4; rm -f stills4/*.jpg
put(){ n=$(python3 -c "print(round(($3+3)*15))"); cp $(printf "%s/f_%05d.jpg" $2 $n) stills4/$1.jpg; }
put A00_현재상태_항공 f4s -2
put A01_베이스_클로즈업 f4c 3.6
put A02_1번반전_항공 f4s 8
put A03_1번반전_클로즈업 f4c 8.5
put A04_1번이동_항공 f4s 18
put A05_2번반전_클로즈업 f4c 29
put A06_2번이동_항공 f4s 39
put A07_안착완료_클로즈업 f4c 42.8
put A08_사상_클로즈업 f4c 45.5
put A09_밀착_클로즈업 f4c 50
put A10_용접_클로즈업 f4c 55
put A11_완성품이동_항공 f4s 65
put A12_완료_항공 f4s 76
put A13_완료_클로즈업 f4c 76
ls stills4 | wc -l
