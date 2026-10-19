#!/bin/bash
cd "$(dirname "$0")"
for v in side closeup; do o=f4${v:0:1}; for k in 0 1 2 3; do START=$((k*300)) END=$(((k+1)*300)) node shot4.mjs $o 15 "view=$v" > r4_${v}_$k.log 2>&1 & done; done
wait; ls f4s | wc -l; ls f4c | wc -l
