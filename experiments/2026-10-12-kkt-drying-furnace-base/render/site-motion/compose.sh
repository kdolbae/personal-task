#!/bin/bash
# usage: compose.sh A|B  -> X_final.mp4 (PiP → 세부 → 우측 상부뷰)
set -e
X=$1; S=f${X}s; C=f${X}c; ENC="-c:v libx264 -crf 20 -preset medium -pix_fmt yuv420p -r 15"
FONT=/root/.fonts/Pretendard-SemiBold.otf
ffmpeg -y -loglevel error -framerate 15 -i $S/f_%05d.jpg -framerate 15 -i $C/f_%05d.jpg -filter_complex \
 "[1:v]crop=1920:926:0:98,scale=640:309[in];[0:v]drawbox=x=18:y=660:w=652:h=355:color=0x0F2A44@0.94:t=fill,drawtext=fontfile=$FONT:text='세부 화면':x=32:y=667:fontsize=21:fontcolor=white[bg];[bg][in]overlay=24:700" \
 $ENC ${X}_pip.mp4
ffmpeg -y -loglevel error -framerate 15 -i $C/f_%05d.jpg $ENC ${X}_cu.mp4
ffmpeg -y -loglevel error -framerate 15 -i $S/f_%05d.jpg $ENC ${X}_side.mp4
printf "file '%s'\n" ${X}_pip.mp4 ${X}_cu.mp4 ${X}_side.mp4 > c${X}.txt
ffmpeg -y -loglevel error -f concat -safe 0 -i c${X}.txt -c copy ${X}_final.mp4
ffprobe -v error -show_entries format=duration -of csv=p=0 ${X}_final.mp4
