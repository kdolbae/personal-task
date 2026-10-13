#!/bin/bash
# usage: compose.sh A|B  -> X_final.mp4 (PiP → 세부 → 우측 상부뷰)
set -e
X=$1; S=f${X}s; C=f${X}c; ENC="-c:v libx264 -crf 20 -preset medium -pix_fmt yuv420p -r 15"
FONT=/root/.fonts/Pretendard-SemiBold.otf
ffmpeg -y -loglevel error -framerate 15 -i $S/f_%05d.jpg -framerate 15 -i $C/f_%05d.jpg -filter_complex \
 "[1:v]crop=1482:926:219:98,scale=640:400,drawtext=fontfile=$FONT:text='세부 화면':x=(w-text_w)/2:y=10:fontsize=21:fontcolor=white:box=1:boxcolor=0x0F1A23@0.86:boxborderw=8[in];[0:v]drawbox=x=22:y=660:w=648:h=408:color=0x10222D@1:t=fill[bg];[bg][in]overlay=26:664" \
 $ENC ${X}_pip.mp4
ffmpeg -y -loglevel error -framerate 15 -i $C/f_%05d.jpg $ENC ${X}_cu.mp4
ffmpeg -y -loglevel error -framerate 15 -i $S/f_%05d.jpg $ENC ${X}_side.mp4
printf "file '%s'\n" ${X}_pip.mp4 ${X}_cu.mp4 ${X}_side.mp4 > c${X}.txt
ffmpeg -y -loglevel error -f concat -safe 0 -i c${X}.txt -c copy ${X}_final.mp4
ffprobe -v error -show_entries format=duration -of csv=p=0 ${X}_final.mp4
